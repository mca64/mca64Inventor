using System.Windows.Forms;

namespace mca64Inventor
{
    public class MainForm : Form
    {
        public MainForm()
        {
            this.Text = "mca64launcher";
            this.Width = 800;
            this.Height = 400;
            var label = new Label { Text = "To jest przyk³adowa forma64.", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleCenter };
            this.Controls.Add(label);
        }
    }
}
