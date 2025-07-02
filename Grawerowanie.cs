using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Inventor;

namespace mca64Inventor
{
    /// <summary>
    /// Klasa logiki grawerowania, wywo�ywana z przycisku na formie.
    /// </summary>
    public class Grawerowanie
    {
        /// <summary>
        /// Wykonuje logik� grawerowania na aktywnym zespole w Inventorze.
        /// </summary>
        public void DoSomething()
        {
            // Pobierz referencj� do aplikacji Inventor
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja == null)
            {
                MessageBox.Show("Nie mo�na uzyska� referencji do Inventora!", "B��d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // Sprawd�, czy otwarty jest dokument zespo�u
            if (aplikacja.ActiveDocument == null)
            {
                MessageBox.Show("Otw�rz dokument zespo�u przed uruchomieniem skryptu!", "B��d", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (!(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
            {
                MessageBox.Show("Aktywny dokument nie jest dokumentem zespo�u!", "B��d", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            // Rozpocznij przetwarzanie zespo�u
            PrzetworzZespol(dokumentZespolu, aplikacja);
        }

        /// <summary>
        /// Przetwarza zesp�: czy�ci folder IGES, graweruje cz�ci i eksportuje do IGES.
        /// </summary>
        private void PrzetworzZespol(AssemblyDocument dokumentZespolu, Inventor.Application aplikacja)
        {
            // Utw�rz �cie�k� do folderu IGES w katalogu zespo�u
            string folderIGES = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dokumentZespolu.FullFileName), "IGES");
            WyczyscFolder(folderIGES, "*.igs");
            var listaCzesci = new List<string>();
            var przetworzonePliki = new List<string>();
            var otwarteDokumenty = new List<Document>();
            // Przetw�rz wszystkie wyst�pienia w zespole
            PrzetworzWystapienia(dokumentZespolu.ComponentDefinition.Occurrences, listaCzesci, folderIGES, przetworzonePliki, otwarteDokumenty, aplikacja);
            // Wy�wietl list� przetworzonych cz�ci
            MessageBox.Show(string.Join(System.Environment.NewLine, listaCzesci), "Lista cz�ci zespo�u", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Czy�ci folder z plikami IGES.
        /// </summary>
        private void WyczyscFolder(string sciezkaFolderu, string maskaPliku)
        {
            if (System.IO.Directory.Exists(sciezkaFolderu))
            {
                foreach (var plik in System.IO.Directory.GetFiles(sciezkaFolderu, maskaPliku))
                {
                    System.IO.File.Delete(plik);
                }
            }
            else
            {
                System.IO.Directory.CreateDirectory(sciezkaFolderu);
            }
        }

        /// <summary>
        /// Przetwarza wyst�pienia w zespole, wykonuje grawerowanie i eksportuje do IGES.
        /// </summary>
        private void PrzetworzWystapienia(ComponentOccurrences wystapienia, List<string> listaCzesci, string folderIGES, List<string> przetworzonePliki, List<Document> otwarteDokumenty, Inventor.Application aplikacja)
        {
            foreach (ComponentOccurrence wystapienie in wystapienia)
            {
                Document dokument = null;
                try { dokument = (Document)wystapienie.Definition.Document; } catch { }
                if (dokument == null) continue;

                string sciezkaPliku = dokument.FullFileName;
                // Rekurencyjnie przetwarzaj podzespo�y
                if (!sciezkaPliku.ToLower().EndsWith(".ipt"))
                {
                    if (dokument is AssemblyDocument assemblyDoc)
                    {
                        PrzetworzWystapienia(assemblyDoc.ComponentDefinition.Occurrences, listaCzesci, folderIGES, przetworzonePliki, otwarteDokumenty, aplikacja);
                    }
                    continue;
                }
                // Pomijaj ju� przetworzone pliki
                if (przetworzonePliki.Contains(sciezkaPliku)) continue;

                // Pobierz tekst grawerowania z w�a�ciwo�ci u�ytkownika
                string tekstGrawerowania = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
                if (string.IsNullOrEmpty(tekstGrawerowania)) continue;

                // Dodaj nazw� wyst�pienia do listy
                if (!listaCzesci.Contains(wystapienie.Name)) listaCzesci.Add(wystapienie.Name);

                // Wykonaj grawerowanie i eksport do IGES
                WykonajGrawerowanie(sciezkaPliku, tekstGrawerowania, przetworzonePliki, otwarteDokumenty, aplikacja);
                EksportujDoIGES(sciezkaPliku, folderIGES, aplikacja);
                przetworzonePliki.Add(sciezkaPliku);
            }
        }

        /// <summary>
        /// Pobiera warto�� w�a�ciwo�ci u�ytkownika z dokumentu.
        /// </summary>
        private string PobierzWlasciwoscUzytkownika(Document dokument, string nazwaWlasciwosci)
        {
            try
            {
                var zestawWlasciwosci = dokument.PropertySets["Inventor User Defined Properties"];
                var wlasciwosc = zestawWlasciwosci[nazwaWlasciwosci] as Inventor.Property;
                return wlasciwosc?.Value?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Wykonuje grawerowanie na pliku cz�ci, je�li posiada w�a�ciwo�� "Grawer".
        /// </summary>
        private void WykonajGrawerowanie(string sciezkaPliku, string tekstGrawerowania, List<string> przetworzonePliki, List<Document> otwarteDokumenty, Inventor.Application aplikacja)
        {
            if (string.IsNullOrEmpty(tekstGrawerowania)) return;
            var dokumentCzesci = aplikacja.Documents.Open(sciezkaPliku) as PartDocument;
            if (dokumentCzesci == null) return;

            // Dodaj szkic i pole tekstowe na p�aszczy�nie YZ
            PlanarSketch szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[2]);
            Point2d pozycjaTekstu = aplikacja.TransientGeometry.CreatePoint2d(-1, 0);
            var poleTekstowe = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
            poleTekstowe.FormattedText = "<StyleOverride FontSize='1.0'>" + tekstGrawerowania + "</StyleOverride>";

            // Usu� poprzednie grawerowanie je�li istnieje
            var funkcjeEkstruzji = dokumentCzesci.ComponentDefinition.Features.ExtrudeFeatures;
            foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
            {
                if (operacjaEkstrudowania.Name == "Grawerowanie64")
                {
                    operacjaEkstrudowania.Delete();
                    break;
                }
            }

            // Dodaj now� operacj� grawerowania (wyci�cie tekstu)
            Profile profil = szkic.Profiles.AddForSolid();
            ExtrudeFeature operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                profil,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kCutOperation
            );
            operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";

            // Je�li grawerowanie nie powiod�o si�, spr�buj na innej p�aszczy�nie
            if ((int)operacjaEkstrudowaniaNowa.HealthStatus == 11780) // kDriverLostHealth
            {
                MessageBox.Show("Nieodpowiednia p�aszczyzna szkicu! Spr�buj� jeszcze raz z inn� p�aszczyzn�.", "B��d ekstruzji", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
                {
                    if (operacjaEkstrudowania.Name == "Grawerowanie64")
                    {
                        operacjaEkstrudowania.Delete();
                        break;
                    }
                }
                szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[3]);
                var poleTekstowe2 = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
                poleTekstowe2.FormattedText = "<StyleOverride FontSize='1.0'>" + tekstGrawerowania + "</StyleOverride>";
                Profile nowyProfil = szkic.Profiles.AddForSolid();
                operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                    nowyProfil,
                    PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                    PartFeatureOperationEnum.kCutOperation
                );
                operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";
            }

            // Zapisz zmiany w pliku cz�ci
            dokumentCzesci.Save();
            otwarteDokumenty.Add(dokumentCzesci as Document);
            przetworzonePliki.Add(sciezkaPliku);
        }

        /// <summary>
        /// Eksportuje plik cz�ci do formatu IGES do wskazanego folderu.
        /// </summary>
        private void EksportujDoIGES(string sciezkaPliku, string folderIGES, Inventor.Application aplikacja)
        {
            if (!System.IO.File.Exists(sciezkaPliku)) return;
            var dokument = aplikacja.Documents.Open(sciezkaPliku);
            if (dokument == null) return;
            string tekstGrawerowania = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
            if (string.IsNullOrEmpty(tekstGrawerowania)) return;
            var tlumaczIGES = aplikacja.ApplicationAddIns.ItemById["{90AF7F44-0C01-11D5-8E83-0010B541CD80}"] as TranslatorAddIn;
            if (tlumaczIGES == null) return;
            var kontekstTlumaczenia = aplikacja.TransientObjects.CreateTranslationContext();
            kontekstTlumaczenia.Type = IOMechanismEnum.kFileBrowseIOMechanism;
        }
    }
}
