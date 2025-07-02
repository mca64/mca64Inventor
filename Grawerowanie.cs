using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Inventor;

namespace mca64Inventor
{
    /// <summary>
    /// Klasa logiki grawerowania, wywo³ywana z przycisku na formie.
    /// </summary>
    public class Grawerowanie
    {
        /// <summary>
        /// Wykonuje logikê grawerowania na aktywnym zespole w Inventorze.
        /// </summary>
        public void DoSomething()
        {
            // Pobierz referencjê do aplikacji Inventor
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja == null)
            {
                MessageBox.Show("Nie mo¿na uzyskaæ referencji do Inventora!", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // SprawdŸ, czy otwarty jest dokument zespo³u
            if (aplikacja.ActiveDocument == null)
            {
                MessageBox.Show("Otwórz dokument zespo³u przed uruchomieniem skryptu!", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (!(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
            {
                MessageBox.Show("Aktywny dokument nie jest dokumentem zespo³u!", "B³¹d", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            // Rozpocznij przetwarzanie zespo³u
            PrzetworzZespol(dokumentZespolu, aplikacja);
        }

        /// <summary>
        /// Przetwarza zespó³: czyœci folder IGES, graweruje czêœci i eksportuje do IGES.
        /// </summary>
        private void PrzetworzZespol(AssemblyDocument dokumentZespolu, Inventor.Application aplikacja)
        {
            // Utwórz œcie¿kê do folderu IGES w katalogu zespo³u
            string folderIGES = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dokumentZespolu.FullFileName), "IGES");
            WyczyscFolder(folderIGES, "*.igs");
            var listaCzesci = new List<string>();
            var przetworzonePliki = new List<string>();
            var otwarteDokumenty = new List<Document>();
            // Przetwórz wszystkie wyst¹pienia w zespole
            PrzetworzWystapienia(dokumentZespolu.ComponentDefinition.Occurrences, listaCzesci, folderIGES, przetworzonePliki, otwarteDokumenty, aplikacja);
            // Wyœwietl listê przetworzonych czêœci
            MessageBox.Show(string.Join(System.Environment.NewLine, listaCzesci), "Lista czêœci zespo³u", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Czyœci folder z plikami IGES.
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
        /// Przetwarza wyst¹pienia w zespole, wykonuje grawerowanie i eksportuje do IGES.
        /// </summary>
        private void PrzetworzWystapienia(ComponentOccurrences wystapienia, List<string> listaCzesci, string folderIGES, List<string> przetworzonePliki, List<Document> otwarteDokumenty, Inventor.Application aplikacja)
        {
            foreach (ComponentOccurrence wystapienie in wystapienia)
            {
                Document dokument = null;
                try { dokument = (Document)wystapienie.Definition.Document; } catch { }
                if (dokument == null) continue;

                string sciezkaPliku = dokument.FullFileName;
                // Rekurencyjnie przetwarzaj podzespo³y
                if (!sciezkaPliku.ToLower().EndsWith(".ipt"))
                {
                    if (dokument is AssemblyDocument assemblyDoc)
                    {
                        PrzetworzWystapienia(assemblyDoc.ComponentDefinition.Occurrences, listaCzesci, folderIGES, przetworzonePliki, otwarteDokumenty, aplikacja);
                    }
                    continue;
                }
                // Pomijaj ju¿ przetworzone pliki
                if (przetworzonePliki.Contains(sciezkaPliku)) continue;

                // Pobierz tekst grawerowania z w³aœciwoœci u¿ytkownika
                string tekstGrawerowania = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
                if (string.IsNullOrEmpty(tekstGrawerowania)) continue;

                // Dodaj nazwê wyst¹pienia do listy
                if (!listaCzesci.Contains(wystapienie.Name)) listaCzesci.Add(wystapienie.Name);

                // Wykonaj grawerowanie i eksport do IGES
                WykonajGrawerowanie(sciezkaPliku, tekstGrawerowania, przetworzonePliki, otwarteDokumenty, aplikacja);
                EksportujDoIGES(sciezkaPliku, folderIGES, aplikacja);
                przetworzonePliki.Add(sciezkaPliku);
            }
        }

        /// <summary>
        /// Pobiera wartoœæ w³aœciwoœci u¿ytkownika z dokumentu.
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
        /// Wykonuje grawerowanie na pliku czêœci, jeœli posiada w³aœciwoœæ "Grawer".
        /// </summary>
        private void WykonajGrawerowanie(string sciezkaPliku, string tekstGrawerowania, List<string> przetworzonePliki, List<Document> otwarteDokumenty, Inventor.Application aplikacja)
        {
            if (string.IsNullOrEmpty(tekstGrawerowania)) return;
            var dokumentCzesci = aplikacja.Documents.Open(sciezkaPliku) as PartDocument;
            if (dokumentCzesci == null) return;

            // Dodaj szkic i pole tekstowe na p³aszczyŸnie YZ
            PlanarSketch szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[2]);
            Point2d pozycjaTekstu = aplikacja.TransientGeometry.CreatePoint2d(-1, 0);
            var poleTekstowe = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
            poleTekstowe.FormattedText = "<StyleOverride FontSize='1.0'>" + tekstGrawerowania + "</StyleOverride>";

            // Usuñ poprzednie grawerowanie jeœli istnieje
            var funkcjeEkstruzji = dokumentCzesci.ComponentDefinition.Features.ExtrudeFeatures;
            foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
            {
                if (operacjaEkstrudowania.Name == "Grawerowanie64")
                {
                    operacjaEkstrudowania.Delete();
                    break;
                }
            }

            // Dodaj now¹ operacjê grawerowania (wyciêcie tekstu)
            Profile profil = szkic.Profiles.AddForSolid();
            ExtrudeFeature operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                profil,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kCutOperation
            );
            operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";

            // Jeœli grawerowanie nie powiod³o siê, spróbuj na innej p³aszczyŸnie
            if ((int)operacjaEkstrudowaniaNowa.HealthStatus == 11780) // kDriverLostHealth
            {
                MessageBox.Show("Nieodpowiednia p³aszczyzna szkicu! Spróbujê jeszcze raz z inn¹ p³aszczyzn¹.", "B³¹d ekstruzji", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

            // Zapisz zmiany w pliku czêœci
            dokumentCzesci.Save();
            otwarteDokumenty.Add(dokumentCzesci as Document);
            przetworzonePliki.Add(sciezkaPliku);
        }

        /// <summary>
        /// Eksportuje plik czêœci do formatu IGES do wskazanego folderu.
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
