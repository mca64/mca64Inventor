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
using WinFormsApplication = System.Windows.Forms.Application;
using SysEnvironment = System.Environment;
using System.Globalization; // Added for CultureInfo

namespace mca64Inventor
{
    public static class TranslationManager
    {
        public enum Language
        {
            English,
            Polish
        }

        private static Language currentLanguage = Language.English;
        private static Dictionary<string, Dictionary<string, string>> translations = new Dictionary<string, Dictionary<string, string>>();

        static TranslationManager()
        {
            // English translations
            translations["English"] = new Dictionary<string, string>
            {
                {"MainFormTitle", "mca64Inventor"},
                {"FontSizeLabel", "Font size:"},
                {"ClosePartsCheckbox", "Close parts after engraving"},
                {"EngraveButton", "Engrave"},
                {"ThumbnailsButton", "Thumbnails"},
                {"PartNameColumn", "Part name"},
                {"PreviewColumn", "Preview"},
                {"EngravingColumn", "mca64"},
                {"NewEngravingColumn", "New Engraving"},
                {"Error", "Error"}, // Generic error title for MessageBox
                {"ErrorInventorNotRunning", "Cannot get reference to Inventor!"},
                {"ErrorOpenAssembly", "Open an assembly document before running the script!"},
                {"ErrorNotAssembly", "The active document is not an assembly document!"},
                {"ErrorSaveAssembly", "Please save the assembly document before running the script!"},
                {"LogFontSizeSaved", "[SETTINGS] Font size saved on window closing: {0}"},
                {"LogFontSizeInvalid", "[SETTINGS] Invalid font size format on window closing. Defaulted to 1.0"},
                {"LogClosePartsSaved", "[SETTINGS] Close parts setting saved: {0}"},
                {"LogThumbnailsGenerated", "Thumbnails have been generated (if possible)"},
                {"LogErrorLoadingParts", "Error loading parts list: {0}"},
                {"LogErrorGeneratingThumbnails", "Error generating thumbnails: {0}"},
                {"ErrorOpeningPartForEngraving", "Error opening part for engraving: {0} - {1}"},
                {"ErrorPartNoComponentDefinition", "Error: Part '{0}' does not have a valid component definition."},
                {"ErrorPartFewerThan3WorkPlanes", "Error: Part '{0}' has fewer than 3 work planes."},
                {"ErrorCouldNotCreateSketch", "Error: Could not create sketch on work plane {0} for part '{1}'."},
                {"ErrorCouldNotCreateTextBox", "Error: Could not create text box for part '{0}'."},
                {"ErrorCouldNotCreateProfile", "Error: Could not create profile for part '{0}'."},
                {"ErrorCouldNotCreateExtrusion", "Error: Could not create extrusion for part '{0}'."},
                {"LogWrongSketchPlane", "Wrong sketch plane! Trying again with another plane."},
                {"LogEngravedParts", "Engraved parts:"},
                {"LogErrorClosingPart", "Error closing part: {0} - {1}"},
                {"LogErrorReturningToAssembly", "Error returning to assembly tab: {0}"},
                {"LogFontSizeUsed", "[ENGRAVING] Font size used from comboBoxFontSize: {0}"},
                {"TooltipEngraveButton", "Launches the engraving form."},
                {"DescriptionEngraveButton", "Button for engraving"},
                {"ErrorInventorNotRunningMsgBox", "Cannot get reference to Inventor!"},
                {"ErrorOpenAssemblyMsgBox", "Open an assembly document before running the script!"},
                {"ErrorNotAssemblyMsgBox", "The active document is not an assembly document!"},
                {"ErrorSaveAssemblyMsgBox", "Please save the assembly document before running the script!"},
                {"ErrorUnexpected", "An unexpected error occurred. \nDetails have been written to the log file located at:\n{0}"},
                {"ErrorAssemblyMissingComponentDefinition", "The active assembly is missing its core component definition and cannot be processed."},
                {"ErrorFileDoesNotExist", "File does not exist: {0}"},
                {"ErrorOpeningPart", "Error opening part: {0} - {1}"},
                {"ErrorGeneratingThumbnail", "Error generating thumbnail for: {0} - {1}"},
                {"ErrorSettingUserProperty", "Error setting user property for '{0}': {1}"},
                {"DebugStyleOverride", "DEBUG: StyleOverride: {0}"},
                {"ErrorSettingCamera", "Error setting camera view for {0}: {1}"}
            };

            // Polish translations
            translations["Polish"] = new Dictionary<string, string>
            {
                {"MainFormTitle", "mca64Inventor"},
                {"FontSizeLabel", "Rozmiar czcionki:"},
                {"ClosePartsCheckbox", "Zamykaj części po grawerowaniu"},
                {"EngraveButton", "Graweruj"},
                {"ThumbnailsButton", "Miniatury"},
                {"PartNameColumn", "Nazwa części"},
                {"PreviewColumn", "Podgląd"},
                {"EngravingColumn", "mca64"},
                {"NewEngravingColumn", "Nowe grawerowanie"},
                {"Error", "Błąd"}, // Generic error title for MessageBox
                {"ErrorInventorNotRunning", "Nie można uzyskać referencji do Inventora!"},
                {"ErrorOpenAssembly", "Otwórz dokument złożenia przed uruchomieniem skryptu!"},
                {"ErrorNotAssembly", "Aktywny dokument nie jest dokumentem złożenia!"},
                {"ErrorSaveAssembly", "Zapisz dokument złożenia przed uruchomieniem skryptu!"},
                {"LogFontSizeSaved", "[USTAWIENIA] Zapisano rozmiar czcionki przy zamykaniu okna: {0}"},
                {"LogFontSizeInvalid", "[USTAWIENIA] Niepoprawny format rozmiaru czcionki przy zamykaniu okna. Ustawiono domyślnie 1.0"},
                {"LogClosePartsSaved", "[USTAWIENIA] Zapisano zamykanie części: {0}"},
                {"LogThumbnailsGenerated", "Miniatury zostały wygenerowane (jeśli to możliwe)"},
                {"LogErrorLoadingParts", "Błąd ładowania listy części: {0}"},
                {"LogErrorGeneratingThumbnails", "Błąd generowania miniatur: {0}"},
                {"ErrorOpeningPartForEngraving", "Błąd otwierania części do grawerowania: {0} - {1}"},
                {"ErrorPartNoComponentDefinition", "Błąd: Część '{0}' nie ma prawidłowej definicji komponentu."},
                {"ErrorPartFewerThan3WorkPlanes", "Błąd: Część '{0}' ma mniej niż 3 płaszczyzny robocze."},
                {"ErrorCouldNotCreateSketch", "Błąd: Nie można utworzyć szkicu na płaszczyźnie roboczej {0} dla części '{1}'."},
                {"ErrorCouldNotCreateTextBox", "Błąd: Nie można utworzyć pola tekstowego dla części '{0}'."},
                {"ErrorCouldNotCreateProfile", "Błąd: Nie można utworzyć profilu dla części '{0}'."},
                {"ErrorCouldNotCreateExtrusion", "Błąd: Nie można utworzyć wyciągnięcia dla części '{0}'."},
                {"LogWrongSketchPlane", "Błędna płaszczyzna szkicu! Ponowna próba z inną płaszczyzną."},
                {"LogEngravedParts", "Wygrawerowane części:"},
                {"LogErrorClosingPart", "Błąd zamykania części: {0} - {1}"},
                {"LogErrorReturningToAssembly", "Błąd podczas powrotu do zakładki ze złożeniem: {0}"},
                {"LogFontSizeUsed", "[GRAWEROWANIE] Użyto rozmiaru czcionki z comboBoxFontSize: {0}"},
                {"TooltipEngraveButton", "Uruchamia formularz grawerowania."},
                {"DescriptionEngraveButton", "Przycisk do grawerowania"},
                {"ErrorInventorNotRunningMsgBox", "Nie można uzyskać referencji do Inventora!"},
                {"ErrorOpenAssemblyMsgBox", "Otwórz dokument złożenia przed uruchomieniem skryptu!"},
                {"ErrorNotAssemblyMsgBox", "Aktywny dokument nie jest dokumentem złożenia!"},
                {"ErrorSaveAssemblyMsgBox", "Zapisz dokument złożenia przed uruchomieniem skryptu!"},
                {"ErrorUnexpected", "Wystąpił nieoczekiwany błąd. \nSzczegóły zostały zapisane w pliku dziennika znajdującym się pod adresem:\n{0}"},
                {"ErrorAssemblyMissingComponentDefinition", "Aktywne złożenie nie posiada definicji komponentu i nie może być przetworzone."},
                {"ErrorFileDoesNotExist", "Plik nie istnieje: {0}"},
                {"ErrorOpeningPart", "Błąd otwierania części: {0} - {1}"},
                {"ErrorGeneratingThumbnail", "Błąd generowania miniatury dla: {0} - {1}"},
                {"ErrorSettingUserProperty", "Błąd ustawiania właściwości użytkownika dla '{0}': {1}"},
                {"DebugStyleOverride", "DEBUG: StyleOverride: {0}"},
                {"ErrorSettingCamera", "Błąd ustawiania widoku kamery dla {0}: {1}"}
            };

            // Set default language based on OS culture
            if (CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "pl")
            {
                currentLanguage = Language.Polish;
            }
            else
            {
                currentLanguage = Language.English;
            }
        }

        public static string GetTranslation(string key, params object[] args)
        {
            if (translations.ContainsKey(currentLanguage.ToString()) && translations[currentLanguage.ToString()].ContainsKey(key))
            {
                return string.Format(translations[currentLanguage.ToString()][key], args);
            }
            // Fallback to English if translation not found
            if (translations["English"].ContainsKey(key))
            {
                return string.Format(translations["English"][key], args);
            }
            return key; // Return key if no translation found
        }

        public static void SetLanguage(Language lang)
        {
            currentLanguage = lang;
        }

        public static Language GetCurrentLanguage()
        {
            return currentLanguage;
        }
    }

    public partial class MainForm : Form
    {
        private static Dictionary<string, int> fontIndexPerAssembly = new Dictionary<string, int>();
        private static Dictionary<string, int> fontSizeIndexPerAssembly = new Dictionary<string, int>();
        private static Dictionary<string, List<(string, Image, string)>> miniaturyPerAssembly = new Dictionary<string, List<(string, Image, string)>>();
        private static Dictionary<string, string> logPerAssembly = new Dictionary<string, string>();
        private static Dictionary<string, MainForm> openFormsPerAssembly = new Dictionary<string, MainForm>();
        private static Dictionary<string, bool> closePartsPerAssembly = new Dictionary<string, bool>();
        private string currentAssemblyPath = null;
        private ComboBox comboBoxLanguage; // Added for language selection

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

        public string SelectedFontName
        {
            get
            {
                if (this.Controls["comboBoxFonts"] is ComboBox cb && cb.SelectedItem is string selectedFont)
                    return selectedFont;
                return "Tahoma"; // Default font
            }
        }

        public MainForm()
        {
            InitializeComponent();
            AddFontComboBox();
            PrepareDataGridView(); // Moved before AddLanguageComboBox
            AddLanguageComboBox(); // Added language combobox
            // Dodaj obsługę comboBoxFontSize
            if (this.Controls["comboBoxFontSize"] is ComboBox cbFontSize)
            {
                cbFontSize.SelectedIndexChanged += ComboBoxFontSize_SelectedIndexChanged;
            }
            // Odczytaj zapisaną wartość rozmiaru czcionki
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
            if (dataGridViewParts.Rows.Count == 0)
                LoadPartsList();
            // Handle immediate engraving update after editing
            dataGridViewParts.CellEndEdit += DataGridViewParts_CellEndEdit;
            dataGridViewParts.EditingControlShowing += DataGridViewParts_EditingControlShowing;
            comboBoxFonts.SelectedIndexChanged += ComboBoxFonts_SelectedIndexChanged;
            this.FormClosing += MainForm_FormClosing;
            UpdateUIStrings(); // Update UI strings after initialization
        }

        public MainForm(string assemblyPath) : this()
        {
            // Ustaw ścieżkę złożenia na starcie
            currentAssemblyPath = assemblyPath;
            // Przywróć stan checkBoxCloseParts dla tego złożenia
            if (this.Controls["checkBoxCloseParts"] is CheckBox cbClose)
            {
                if (closePartsPerAssembly.TryGetValue(assemblyPath, out bool val))
                    cbClose.Checked = val;
                else if (Properties.Settings.Default["ClosePartsAfterEngraving_" + assemblyPath] != null)
                    cbClose.Checked = (bool)Properties.Settings.Default["ClosePartsAfterEngraving_" + assemblyPath];
                else
                    cbClose.Checked = Properties.Settings.Default.ClosePartsAfterEngraving;
            }
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

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Przywróć ustawienie zamykania części
            if (this.Controls["checkBoxCloseParts"] is CheckBox cbClose)
            {
                cbClose.Checked = Properties.Settings.Default.ClosePartsAfterEngraving;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Controls["comboBoxFontSize"] is ComboBox cb)
            {
                if (cb.SelectedItem is string selected && float.TryParse(selected.Replace(',', '.'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float parsed) && parsed > 0)
                {
                    Properties.Settings.Default.FontSize = parsed;
                    LogMessage(TranslationManager.GetTranslation("LogFontSizeSaved", parsed));
                }
                else
                {
                    cb.SelectedIndex = 0;
                    Properties.Settings.Default.FontSize = 1.0f;
                    LogMessage(TranslationManager.GetTranslation("LogFontSizeInvalid"));
                }
            }
            // Zapisz ustawienie zamykania części
            if (this.Controls["checkBoxCloseParts"] is CheckBox cbClose)
            {
                Properties.Settings.Default.ClosePartsAfterEngraving = cbClose.Checked;
                if (!string.IsNullOrEmpty(currentAssemblyPath))
                {
                    closePartsPerAssembly[currentAssemblyPath] = cbClose.Checked;

                    // Poprawiona logika zapisu indywidualnego dla złożenia
                    string settingName = "ClosePartsAfterEngraving_" + currentAssemblyPath.Replace("\\", "_").Replace(":", "");
                    bool settingExists = false;
                    foreach (SettingsProperty property in Properties.Settings.Default.Properties)
                    {
                        if (property.Name == settingName)
                        {
                            settingExists = true;
                            break;
                        }
                    }

                    if (!settingExists)
                    {
                        // Utwórz nowe ustawienie, jeśli nie istnieje
                        SettingsProperty newSetting = new SettingsProperty(settingName);
                        newSetting.PropertyType = typeof(bool);
                        newSetting.DefaultValue = false;
                        newSetting.Provider = Properties.Settings.Default.Providers["LocalFileSettingsProvider"];
                        newSetting.Attributes.Add(typeof(System.Configuration.UserScopedSettingAttribute), new System.Configuration.UserScopedSettingAttribute());
                        Properties.Settings.Default.Properties.Add(newSetting);
                        Properties.Settings.Default.Reload(); // Przeładuj, aby nowe ustawienie było dostępne
                    }

                    Properties.Settings.Default[settingName] = cbClose.Checked;
                }
                LogMessage(TranslationManager.GetTranslation("LogClosePartsSaved", cbClose.Checked));
            }
            Properties.Settings.Default.Save();
        }

        private void PrepareDataGridView()
        {
            dataGridViewParts.Columns.Clear();
            dataGridViewParts.RowTemplate.Height = 64;
            dataGridViewParts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewParts.Columns.Add("PartName", TranslationManager.GetTranslation("PartNameColumn"));
            var imageCol = new DataGridViewImageColumn();
            imageCol.Name = "Preview";
            imageCol.HeaderText = TranslationManager.GetTranslation("PreviewColumn");
            imageCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
            dataGridViewParts.Columns.Add(imageCol);
            var grawerCol = new DataGridViewTextBoxColumn();
            grawerCol.Name = "Grawer";
            grawerCol.HeaderText = TranslationManager.GetTranslation("EngravingColumn");
            grawerCol.ReadOnly = true;
            dataGridViewParts.Columns.Add(grawerCol);
            // Column for editing engraving as DataGridViewTextBoxColumn
            var editCol = new DataGridViewTextBoxColumn();
            editCol.Name = "EditGrawer";
            editCol.HeaderText = TranslationManager.GetTranslation("NewEngravingColumn");
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

        private void AddLanguageComboBox()
        {
            comboBoxLanguage = new ComboBox();
            comboBoxLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxLanguage.Location = new System.Drawing.Point(10, 10); // Top-left
            comboBoxLanguage.Width = 100;
            comboBoxLanguage.Name = "comboBoxLanguage";
            comboBoxLanguage.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            comboBoxLanguage.Items.Add("English");
            comboBoxLanguage.Items.Add("Polski");
            // Temporarily unsubscribe to prevent premature UpdateUIStrings call
            comboBoxLanguage.SelectedIndexChanged -= ComboBoxLanguage_SelectedIndexChanged;

            if (TranslationManager.GetCurrentLanguage() == TranslationManager.Language.Polish)
            {
                comboBoxLanguage.SelectedItem = "Polski";
            }
            else
            {
                comboBoxLanguage.SelectedItem = "English";
            }

            // Re-subscribe after setting the selected item
            comboBoxLanguage.SelectedIndexChanged += ComboBoxLanguage_SelectedIndexChanged;

            this.Controls.Add(comboBoxLanguage);
            comboBoxLanguage.BringToFront();
        }

        private void ComboBoxLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxLanguage.SelectedItem.ToString() == "Polski")
            {
                TranslationManager.SetLanguage(TranslationManager.Language.Polish);
            }
            else
            {
                TranslationManager.SetLanguage(TranslationManager.Language.English);
            }
            UpdateUIStrings();
        }

        private void UpdateUIStrings()
        {
            Version assemblyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string versionString = $"v{assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.Build}.{assemblyVersion.Revision}";
            this.Text = TranslationManager.GetTranslation("MainFormTitle") + " " + versionString;
            labelFontSize.Text = TranslationManager.GetTranslation("FontSizeLabel");
            checkBoxCloseParts.Text = TranslationManager.GetTranslation("ClosePartsCheckbox");
            buttonGrawerowanie.Text = TranslationManager.GetTranslation("EngraveButton");
            buttonGenerateThumbnails.Text = TranslationManager.GetTranslation("ThumbnailsButton");

            dataGridViewParts.Columns["PartName"].HeaderText = TranslationManager.GetTranslation("PartNameColumn");
            dataGridViewParts.Columns["Preview"].HeaderText = TranslationManager.GetTranslation("PreviewColumn");
            dataGridViewParts.Columns["Grawer"].HeaderText = TranslationManager.GetTranslation("EngravingColumn");
            dataGridViewParts.Columns["EditGrawer"].HeaderText = TranslationManager.GetTranslation("NewEngravingColumn");
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
            var czesci = mca64Inventor.CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja, false);
            var czesc = czesci.Find(c => c.Nazwa == partName);
            if (czesc == null) return;
            if ((czesc.Grawer ?? string.Empty) == (newGrawer ?? string.Empty))
                return;
            bool ok = CzesciHelper.UstawWlasciwoscUzytkownika(czesc.SciezkaPliku, "Grawer", newGrawer, aplikacja, LogMessage);
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
                var czesci = mca64Inventor.CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja, false);
                dataGridViewParts.Rows.Clear();
                foreach (var czesc in czesci)
                {
                    // If thumbnail is null, set empty bitmap 1x1 to avoid display error
                    var img = czesc.Miniatura ?? new Bitmap(1, 1);
                    dataGridViewParts.Rows.Add(czesc.Nazwa, img, czesc.Grawer, "");
                }
                LogMessage(TranslationManager.GetTranslation("LogThumbnailsGenerated"));
                // After generating thumbnails, save them to memory for the given assembly
                if (!string.IsNullOrEmpty(currentAssemblyPath))
                {
                    var miniatury = new List<(string, Image, string)>();
                    foreach (DataGridViewRow row in dataGridViewParts.Rows)
                    {
                        if (row.Cells[0].Value != null)
                        {
                            Image originalImage = row.Cells[1].Value as Image;
                            Image clonedImage = originalImage != null ? (Image)originalImage.Clone() : new Bitmap(1, 1);
                            miniatury.Add((row.Cells[0].Value.ToString(), clonedImage, row.Cells[2].Value?.ToString()));
                        }
                    }
                    miniaturyPerAssembly[currentAssemblyPath] = miniatury;
                }
            }
            catch (Exception ex)
            {
                LogMessage(TranslationManager.GetTranslation("LogErrorLoadingParts", ex.Message));
            }
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
                var czesci = mca64Inventor.CzesciHelper.GenerujMiniaturyDlaCzesci(dokumentZespolu, aplikacja);
                dataGridViewParts.Rows.Clear();
                foreach (var czesc in czesci)
                {
                    var img = czesc.Miniatura ?? new Bitmap(1, 1);
                    dataGridViewParts.Rows.Add(czesc.Nazwa, img, czesc.Grawer, "");
                }
                LogMessage(TranslationManager.GetTranslation("LogThumbnailsGenerated"));
                // After generating thumbnails, save them to memory for the given assembly
                if (!string.IsNullOrEmpty(currentAssemblyPath))
                {
                    var miniatury = new List<(string, Image, string)>();
                    foreach (DataGridViewRow row in dataGridViewParts.Rows)
                    {
                        if (row.Cells[0].Value != null)
                        {
                            Image originalImage = row.Cells[1].Value as Image;
                            Image clonedImage = originalImage != null ? (Image)originalImage.Clone() : new Bitmap(1, 1);
                            miniatury.Add((row.Cells[0].Value.ToString(), clonedImage, row.Cells[2].Value?.ToString()));
                        }
                    }
                    miniaturyPerAssembly[currentAssemblyPath] = miniatury;
                }
            }
            catch (Exception ex)
            {
                LogMessage(TranslationManager.GetTranslation("LogErrorGeneratingThumbnails", ex.Message));
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
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja == null)
            {
                MessageBox.Show(TranslationManager.GetTranslation("ErrorInventorNotRunningMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogMessage(TranslationManager.GetTranslation("ErrorInventorNotRunning"));
                return;
            }
            if (aplikacja.ActiveDocument == null)
            {
                MessageBox.Show(TranslationManager.GetTranslation("ErrorOpenAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                LogMessage(TranslationManager.GetTranslation("ErrorOpenAssembly"));
                return;
            }
            if (!(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
            {
                MessageBox.Show(TranslationManager.GetTranslation("ErrorNotAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                LogMessage(TranslationManager.GetTranslation("ErrorNotAssembly"));
                return;
            }

            var logic = new mca64Inventor.Grawerowanie();
            bool zamykajCzesci = false;
            if (this.Controls["checkBoxCloseParts"] is CheckBox cb)
                zamykajCzesci = cb.Checked;

            logic.DoSomething(GlobalFontSize, zamykajCzesci);
        }
    }
}
