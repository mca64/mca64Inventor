using System.Windows.Forms; // Importuje klasy do obs³ugi okienek i kontrolek Windows Forms

namespace mca64Inventor // Przestrzeñ nazw grupuj¹ca klasy dodatku
{
    /// <summary>
    /// G³ówna forma (okno) wyœwietlana po klikniêciu przycisku w Inventorze.
    /// </summary>
    public class MainForm : Form
    {
        /// <summary>
        /// Konstruktor ustawia tytu³, rozmiar i dodaje etykietê do okna.
        /// </summary>
        public MainForm()
        {
            // Ustaw tytu³ okna
            this.Text = "mca64launcher";
            // Ustaw szerokoœæ i wysokoœæ okna
            this.Width = 800;
            this.Height = 400;
            // Utwórz etykietê z tekstem i wyœrodkuj j¹
            var label = new Label
            {
                Text = "To jest przyk³adowa forma64.",
                Dock = DockStyle.Fill, // Etykieta zajmuje ca³e okno
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter // Wyœrodkowanie tekstu
            };
            // Dodaj etykietê do okna
            this.Controls.Add(label);
        }
    }
}
