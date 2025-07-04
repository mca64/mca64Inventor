using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Inventor;
using System.Linq;
using System.Configuration;
// Add aliases for ambiguous types
using WinFormsTextBox = System.Windows.Forms.TextBox;
using SysEnvironment = System.Environment;

namespace mca64Inventor
{
    public partial class MainForm : Form
    {
        private float globalFontSize = 1.0f;
        private static Dictionary<string, int> fontIndexPerAssembly = new Dictionary<string, int>();
        private static Dictionary<string, int> fontSizeIndexPerAssembly = new Dictionary<string, int>();
        private static Dictionary<string, List<(string, Image, string)>> miniaturyPerAssembly = new Dictionary<string, List<(string, Image, string)>>();
        private static Dictionary<string, string> logPerAssembly = new Dictionary<string, string>();
        private static Dictionary<string, MainForm> openFormsPerAssembly = new Dictionary<string, MainForm>();
        private string currentAssemblyPath = null;

        public float GlobalFontSize
        {
            get
            {
                float val = 1.0f;
                if (this.Controls["comboBoxFontSize"] is ComboBox cb && cb.SelectedItem is string selected && float.TryParse(selected.Replace(',', '.'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float parsed) && parsed > 0)
                    val = parsed;
                return val;
            }
        }

        public MainForm()
        {
            InitializeComponent();
            AddFontComboBox();
            // Dodaj obs³ugê comboBoxFontSize
            if (this.Controls["comboBoxFontSize"] is ComboBox cbFontSize)
            {
                cbFontSize.SelectedIndexChanged += ComboBoxFontSize_SelectedIndexChanged;
            }
            // Odczytaj zapisan¹ wartoœæ rozmiaru czcionki
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja != null && aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu)
            {
                currentAssemblyPath = dokumentZespolu.FullFileName;
                // Restore font selection
                if (fontIndexPerAssembly.TryGetValue(currentAssemblyPath, out int idx) && idx >= 0 && idx < comboBoxFonts.Items.Count)
                    comboBoxFonts.SelectedIndex = idx;
                // Restore font size selection
                if (fontSizeIndexPerAssembly.TryGetValue(currentAssemblyPath, out int idxSize) && idxSize >= 0 && idxSize < comboBoxFontSize.Items.Count)
                    comboBoxFontSize.SelectedIndex = idxSize;
                // Restore thumbnails
                if (miniaturyPerAssembly.TryGetValue(currentAssemblyPath, out var miniatury))
                {
                    dataGridViewParts.Rows.Clear();
                    foreach (var (nazwa, img, grawer) in miniatury)
                        dataGridViewParts.Rows.Add(nazwa, img, grawer, grawer);
                }
                // Restore log
                if (logPerAssembly.TryGetValue(currentAssemblyPath, out var log))
                    textBoxLog.Text = log;
            }
            // Set form size to 3/4 of the screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Width = (int)(screen.Width * 0.75);
            this.Height = (int)(screen.Height * 0.75);
            this.StartPosition = FormStartPosition.CenterScreen;
            // Prepare DataGridView
            PrepareDataGridView();
            if (dataGridViewParts.Rows.Count == 0)
                LoadPartsList();
            // Handle immediate engraving update after editing
            dataGridViewParts.CellEndEdit += DataGridViewParts_CellEndEdit;
            dataGridViewParts.EditingControlShowing += DataGridViewParts_EditingControlShowing;
            comboBoxFonts.SelectedIndexChanged += ComboBoxFonts_SelectedIndexChanged;
            this.FormClosing += MainForm_FormClosing;
        }

        public MainForm(string assemblyPath) : this()
        {
            // Ustaw œcie¿kê z³o¿enia na starcie
            currentAssemblyPath = assemblyPath;
        }

        public static MainForm ShowForAssembly(string assemblyPath)
        {
            if (string.IsNullOrEmpty(assemblyPath)) return null;
            if (openFormsPerAssembly.TryGetValue(assemblyPath, out var existingForm))
            {
                if (!existingForm.IsDisposed)
                {
                    existingForm.WindowState = FormWindowState.Normal;
                    existingForm.BringToFront();
                    existingForm.Focus();
                    return existingForm;
                }
                else
                {
                    openFormsPerAssembly.Remove(assemblyPath);
                }
            }
            var form = new MainForm(assemblyPath);
            openFormsPerAssembly[assemblyPath] = form;
            form.FormClosed += (s, e) => openFormsPerAssembly.Remove(assemblyPath);
            form.Show();
            return form;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Controls["comboBoxFontSize"] is ComboBox cb)
            {
                if (cb.SelectedItem is string selected && float.TryParse(selected.Replace(',', '.'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float parsed) && parsed > 0)
                {
                    Properties.Settings.Default.FontSize = parsed;
                    Properties.Settings.Default.Save();
                    LogMessage($"[USTAWIENIA] Zapisano rozmiar czcionki przy zamykaniu okna: {parsed}");
                }
                else
                {
                    cb.SelectedIndex = 0;
                    Properties.Settings.Default.FontSize = 1.0f;
                    Properties.Settings.Default.Save();
                    LogMessage("[USTAWIENIA] Niepoprawny format rozmiaru czcionki przy zamykaniu okna. Ustawiono domyœlnie 1.0");
                }
            }
        }

        private void PrepareDataGridView()
        {
            dataGridViewParts.Columns.Clear();
            dataGridViewParts.RowTemplate.Height = 64;
            dataGridViewParts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewParts.Columns.Add("PartName", "Part name");
            var imageCol = new DataGridViewImageColumn();
            imageCol.Name = "Preview";
            imageCol.HeaderText = "Preview";
            imageCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
            dataGridViewParts.Columns.Add(imageCol);
            var grawerCol = new DataGridViewTextBoxColumn();
            grawerCol.Name = "Grawer";
            grawerCol.HeaderText = "Engraving";
            grawerCol.ReadOnly = true;
            dataGridViewParts.Columns.Add(grawerCol);
            // Column for editing engraving as DataGridViewTextBoxColumn
            var editCol = new DataGridViewTextBoxColumn();
            editCol.Name = "EditGrawer";
            editCol.HeaderText = "New Engraving";
            editCol.ReadOnly = false;
            dataGridViewParts.Columns.Add(editCol);
            // Removed button from row
            // One global button below the table
        }

        private void AddFontComboBox()
        {
            if (this.Controls.ContainsKey("comboBoxFonts"))
                return; // Do not add again
            comboBoxFonts = new ComboBox();
            comboBoxFonts.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxFonts.Location = new System.Drawing.Point(440, 10); // next to buttonUpdateAllGrawer
            comboBoxFonts.Width = 200;
            comboBoxFonts.Name = "comboBoxFonts";
            comboBoxFonts.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            // Get list of system fonts
            int tahomaIndex = -1;
            int i = 0;
            foreach (var font in System.Drawing.FontFamily.Families)
            {
                comboBoxFonts.Items.Add(font.Name);
                if (font.Name.Equals("Tahoma", StringComparison.OrdinalIgnoreCase))
                    tahomaIndex = i;
                i++;
            }
            if (tahomaIndex >= 0)
                comboBoxFonts.SelectedIndex = tahomaIndex;
            else if (comboBoxFonts.Items.Count > 0)
                comboBoxFonts.SelectedIndex = 0;
            // Add to form at the very top
            this.Controls.Add(comboBoxFonts);
            comboBoxFonts.BringToFront();

            // Add comboBoxFontSize
            if (!this.Controls.ContainsKey("comboBoxFontSize"))
            {
                comboBoxFontSize = new ComboBox();
                comboBoxFontSize.DropDownStyle = ComboBoxStyle.DropDownList;
                comboBoxFontSize.Location = new System.Drawing.Point(650, 10); // next to comboBoxFonts
                comboBoxFontSize.Width = 100;
                comboBoxFontSize.Name = "comboBoxFontSize";
                comboBoxFontSize.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                // Add font size options
                comboBoxFontSize.Items.Clear();
                var culture = System.Globalization.CultureInfo.CurrentCulture;
                for (double v = 0.5; v <= 10.0; v += 0.5)
                    comboBoxFontSize.Items.Add(v.ToString("0.0", culture));
                comboBoxFontSize.SelectedIndex = 1; // Default to "1,0" or "1.0" depending on culture
                this.Controls.Add(comboBoxFontSize);
                comboBoxFontSize.BringToFront();
            }
        }

        private void ComboBoxFonts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(currentAssemblyPath))
                fontIndexPerAssembly[currentAssemblyPath] = comboBoxFonts.SelectedIndex;
        }

        private void ComboBoxFontSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(currentAssemblyPath) && this.Controls["comboBoxFontSize"] is ComboBox cb)
                fontSizeIndexPerAssembly[currentAssemblyPath] = cb.SelectedIndex;
        }

        private void DataGridViewParts_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Update only the edited row
            var row = dataGridViewParts.Rows[e.RowIndex];
            string partName = row.Cells["PartName"].Value?.ToString();
            string newGrawer = row.Cells["EditGrawer"].Value?.ToString() ?? string.Empty;
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja == null || aplikacja.ActiveDocument == null || !(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
                return;
            var czesci = CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja, false);
            var czesc = czesci.Find(c => c.Nazwa == partName);
            if (czesc == null) return;
            if ((czesc.Grawer ?? string.Empty) == (newGrawer ?? string.Empty))
                return;
            bool ok = CzesciHelper.UstawWlasciwoscUzytkownika(czesc.SciezkaPliku, "Grawer", newGrawer, aplikacja);
            if (ok)
            {
                row.Cells["Grawer"].Value = newGrawer;
            }
            // Do not create a new form, do not close, do not hide!
        }

        private void DataGridViewParts_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is WinFormsTextBox tb)
            {
                tb.KeyDown -= DataGridViewParts_EditGrawer_KeyDown;
                tb.KeyDown += DataGridViewParts_EditGrawer_KeyDown;
                // Set to not select all text upon entering edit mode
                tb.Click -= EditGrawerTextBox_Click;
                tb.Click += EditGrawerTextBox_Click;
                tb.Enter -= EditGrawerTextBox_Enter;
                tb.Enter += EditGrawerTextBox_Enter;
            }
        }

        private void EditGrawerTextBox_Click(object sender, EventArgs e)
        {
            if (sender is WinFormsTextBox tb)
            {
                // If there is no selection, set cursor to the end
                if (tb.SelectionLength == tb.TextLength)
                {
                    tb.SelectionStart = tb.TextLength;
                    tb.SelectionLength = 0;
                }
            }
        }

        private void EditGrawerTextBox_Enter(object sender, EventArgs e)
        {
            if (sender is WinFormsTextBox tb)
            {
                // Set cursor to the end of the text, do not select everything
                tb.SelectionStart = tb.TextLength;
                tb.SelectionLength = 0;
            }
        }

        private void DataGridViewParts_EditGrawer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                if (dataGridViewParts.IsCurrentCellInEditMode)
                {
                    dataGridViewParts.EndEdit();
                }
                // Do not close, do not hide, do not create a new form!
            }
        }

        private void LoadPartsList()
        {
            try
            {
                var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
                if (aplikacja == null || aplikacja.ActiveDocument == null || !(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
                    return;
                var czesci = CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja, false);
                dataGridViewParts.Rows.Clear();
                foreach (var czesc in czesci)
                {
                    // If thumbnail is null, set empty bitmap 1x1 to avoid display error
                    var img = czesc.Miniatura ?? new Bitmap(1, 1);
                    dataGridViewParts.Rows.Add(czesc.Nazwa, img, czesc.Grawer, czesc.Grawer);
                }
            }
            catch { }
        }

        private void buttonGenerateThumbnails_Click(object sender, EventArgs e)
        {
            var prevState = this.WindowState;
            this.WindowState = FormWindowState.Minimized;
            bool prevTopMost = this.TopMost;
            this.TopMost = false;
            try
            {
                var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
                if (aplikacja == null || aplikacja.ActiveDocument == null || !(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
                    return;
                var czesci = CzesciHelper.GenerujMiniaturyDlaCzesci(dokumentZespolu, aplikacja);
                dataGridViewParts.Rows.Clear();
                foreach (var czesc in czesci)
                {
                    var img = czesc.Miniatura ?? new Bitmap(1, 1);
                    dataGridViewParts.Rows.Add(czesc.Nazwa, img, czesc.Grawer, "");
                }
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage("Thumbnails have been generated (if possible)");
                        break;
                    }
                }
                // After generating thumbnails, save them to memory for the given assembly
                if (!string.IsNullOrEmpty(currentAssemblyPath))
                {
                    var miniatury = new List<(string, Image, string)>();
                    foreach (DataGridViewRow row in dataGridViewParts.Rows)
                    {
                        if (row.Cells[0].Value != null)
                            miniatury.Add((row.Cells[0].Value.ToString(), row.Cells[1].Value as Image, row.Cells[2].Value?.ToString()));
                    }
                    miniaturyPerAssembly[currentAssemblyPath] = miniatury;
                }
            }
            catch (Exception ex)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error generating thumbnails: {ex.Message}");
                        break;
                    }
                }
            }
            finally
            {
                this.TopMost = prevTopMost;
                this.WindowState = prevState;
            }
        }

        // Add helper method for logging
        public void LogMessage(string msg)
        {
            if (this.textBoxLog.InvokeRequired)
            {
                this.textBoxLog.Invoke(new Action(() => LogMessage(msg)));
                return;
            }
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string[] lines = msg.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            foreach (var line in lines)
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    string entry = $"[{timestamp}] {line}" + SysEnvironment.NewLine;
                    this.textBoxLog.AppendText(entry);
                    if (!string.IsNullOrEmpty(currentAssemblyPath))
                    {
                        if (!logPerAssembly.ContainsKey(currentAssemblyPath))
                            logPerAssembly[currentAssemblyPath] = "";
                        logPerAssembly[currentAssemblyPath] += entry;
                    }
                }
            }
        }

        // Restore engraving button logic
        private void buttonGrawerowanie_Click(object sender, EventArgs e)
        {
            var logic = new Grawerowanie();
            bool zamykajCzesci = false;
            if (this.Controls["checkBoxCloseParts"] is CheckBox cb)
                zamykajCzesci = cb.Checked;
            logic.DoSomething(GlobalFontSize, zamykajCzesci);
        }
    }
}
