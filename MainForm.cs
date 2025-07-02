using System;
using System.Windows.Forms;

namespace mca64Inventor
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            // Ustaw rozmiar formy na 3/4 ekranu
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Width = (int)(screen.Width * 0.75);
            this.Height = (int)(screen.Height * 0.75);
            this.StartPosition = FormStartPosition.CenterScreen;
            // Ustaw przycisk na samym dole, na ca³¹ szerokoœæ
            buttonGrawerowanie.Width = this.ClientSize.Width - 40;
            buttonGrawerowanie.Left = 20;
            buttonGrawerowanie.Top = this.ClientSize.Height - buttonGrawerowanie.Height - 20;
            buttonGrawerowanie.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
        }

        private void buttonGrawerowanie_Click(object sender, EventArgs e)
        {
            this.Hide(); // Ukryj formê przed rozpoczêciem grawerowania
            var logic = new Grawerowanie();
            logic.DoSomething();
            this.Close(); // Zamknij formê po zakoñczeniu
        }
    }
}
