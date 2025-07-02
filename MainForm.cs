using System.Windows.Forms; // Importuje klasy do obs�ugi okienek i kontrolek Windows Forms

namespace mca64Inventor // Przestrze� nazw grupuj�ca klasy dodatku
{
    /// <summary>
    /// G��wna forma (okno) wy�wietlana po klikni�ciu przycisku w Inventorze.
    /// </summary>
    public class MainForm : Form
    {
        /// <summary>
        /// Konstruktor ustawia tytu�, rozmiar i dodaje etykiet� do okna.
        /// </summary>
        public MainForm()
        {
            // Ustaw tytu� okna
            this.Text = "mca64launcher";
            // Ustaw szeroko�� i wysoko�� okna
            this.Width = 800;
            this.Height = 400;
            // Utw�rz etykiet� z tekstem i wy�rodkuj j�
            var label = new Label
            {
                Text = "To jest przyk�adowa forma64.",
                Dock = DockStyle.Fill, // Etykieta zajmuje ca�e okno
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter // Wy�rodkowanie tekstu
            };
            // Dodaj etykiet� do okna
            this.Controls.Add(label);
        }
    }
}
