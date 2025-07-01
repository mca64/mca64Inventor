using System.Windows.Forms;

namespace mca64Inventor
{
    public class MainForm : Form
    {
        public MainForm()
        {
            this.Text = "mca64launcher";
            this.Width = 400;
            this.Height = 200;
            var label = new Label { Text = "To jest przyk³adowa forma.", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleCenter };
            this.Controls.Add(label);
        }
    }
}
