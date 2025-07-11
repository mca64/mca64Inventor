using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Inventor;
using System.Globalization; // Dodano dla CultureInfo

namespace mca64Inventor
{
    /// <summary>
    /// Klasa modelu danych dla części używanych do wyświetlania i grawerowania.
    /// </summary>
    public class CzesciInfo
    {
        // Nazwa części.
        public string Nazwa { get; set; }
        // Pełna ścieżka do pliku części.
        public string SciezkaPliku { get; set; }
        // Miniatura części (obraz).
        public Image Miniatura { get; set; }
        // Tekst do grawerowania na części.
        public string Grawer { get; set; }
    }

    /// <summary>
    /// Klasa pomocnicza do operacji na częściach złożenia.
    /// </summary>
    public static class CzesciHelper
    {
        /// <summary>
        /// Pobiera listę części ze złożenia wraz z miniaturami i polem grawerowania.
        /// </summary>
        /// <param name="dokumentZespolu">Dokument złożenia Inventor.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        /// <param name="generateThumbnails">Czy generować miniatury (domyślnie false).</param>
        /// <returns>Lista obiektów CzesciInfo.</returns>
        public static List<CzesciInfo> PobierzCzesciZespolu(AssemblyDocument dokumentZespolu, Inventor.Application aplikacja, bool generateThumbnails = false)
        {
            var lista = new List<CzesciInfo>();
            // Sprawdza, czy definicja komponentu złożenia istnieje.
            if (dokumentZespolu.ComponentDefinition == null)
            {
                // Wyświetla komunikat o błędzie, jeśli definicja komponentu jest brakująca.
                MessageBox.Show(TranslationManager.GetTranslation("ErrorAssemblyMissingComponentDefinition"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lista;
            }
            // Przetwarza wystąpienia komponentów rekurencyjnie.
            PrzetworzWystapienia(dokumentZespolu.ComponentDefinition.Occurrences, lista, new HashSet<string>(), aplikacja, generateThumbnails);
            return lista;
        }

        /// <summary>
        /// Rekurencyjnie przetwarza wystąpienia komponentów w złożeniu.
        /// </summary>
        /// <param name="wystapienia">Kolekcja wystąpień komponentów.</param>
        /// <param name="lista">Lista do wypełnienia informacjami o częściach.</param>
        /// <param name="przetworzonePliki">Zbiór ścieżek do już przetworzonych plików, aby uniknąć duplikatów.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        /// <param name="generateThumbnails">Czy generować miniatury.</param>
        private static void PrzetworzWystapienia(ComponentOccurrences wystapienia, List<CzesciInfo> lista, HashSet<string> przetworzonePliki, Inventor.Application aplikacja, bool generateThumbnails)
        {
            // Iteruje przez każde wystąpienie komponentu.
            foreach (ComponentOccurrence wystapienie in wystapienia)
            {
                Document dokument = null;
                try { dokument = (Document)wystapienie.Definition.Document; } catch { } // Próbuje pobrać dokument powiązany z wystąpieniem.
                if (dokument == null) continue; // Pomija, jeśli dokument nie istnieje.

                string sciezkaPliku = dokument.FullFileName;
                // Sprawdza, czy plik nie jest plikiem .ipt (częścią).
                if (!sciezkaPliku.ToLower().EndsWith(".ipt"))
                {
                    // Jeśli to złożenie, przetwarza jego wystąpienia rekurencyjnie.
                    if (dokument is AssemblyDocument assemblyDoc)
                        PrzetworzWystapienia(assemblyDoc.ComponentDefinition.Occurrences, lista, przetworzonePliki, aplikacja, generateThumbnails);
                    continue; // Pomija dalsze przetwarzanie, jeśli to nie część.
                }
                // Sprawdza, czy plik został już przetworzony.
                if (przetworzonePliki.Contains(sciezkaPliku)) continue;
                przetworzonePliki.Add(sciezkaPliku); // Dodaje plik do zbioru przetworzonych.

                // Pobiera niestandardową właściwość "Grawer" z dokumentu.
                string grawer = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
                // Dodaje nową informację o części do listy.
                lista.Add(new CzesciInfo { Nazwa = wystapienie.Name, SciezkaPliku = sciezkaPliku, Miniatura = null, Grawer = grawer });
            }
        }

        /// <summary>
        /// Generuje miniatury dla wszystkich części .ipt w złożeniu. Otwiera każdą część, robi zrzut ekranu,
        /// i zamyka tylko te części, które zostały tymczasowo otwarte.
        /// </summary>
        /// <param name="dokumentZespolu">Dokument złożenia Inventor.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        /// <returns>Lista obiektów CzesciInfo z wygenerowanymi miniaturami.</returns>
        public static List<CzesciInfo> GenerujMiniaturyDlaCzesci(AssemblyDocument dokumentZespolu, Inventor.Application aplikacja)
        {
            var lista = new List<CzesciInfo>();
            var przetworzonePliki = new HashSet<string>();
            // Rozpoczyna rekurencyjne generowanie miniatur.
            GenerujMiniaturyRekurencyjnie(dokumentZespolu.ComponentDefinition.Occurrences, lista, przetworzonePliki, aplikacja);
            return lista;
        }

        /// <summary>
        /// Rekurencyjnie generuje miniatury dla wystąpień komponentów.
        /// </summary>
        /// <param name="wystapienia">Kolekcja wystąpień komponentów.</param>
        /// <param name="lista">Lista do wypełnienia informacjami o częściach i miniaturami.</param>
        /// <param name="przetworzonePliki">Zbiór ścieżek do już przetworzonych plików.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        private static void GenerujMiniaturyRekurencyjnie(ComponentOccurrences wystapienia, List<CzesciInfo> lista, HashSet<string> przetworzonePliki, Inventor.Application aplikacja)
        {
            foreach (ComponentOccurrence wystapienie in wystapienia)
            {
                Document dokument = null;
                try { dokument = (Document)wystapienie.Definition.Document; } catch { }
                if (dokument == null) continue;

                string sciezkaPliku = dokument.FullFileName;
                // Jeśli to podzłożenie, rekurencyjnie generuje miniatury dla jego komponentów.
                if (dokument is AssemblyDocument subAsm)
                {
                    GenerujMiniaturyRekurencyjnie(subAsm.ComponentDefinition.Occurrences, lista, przetworzonePliki, aplikacja);
                    continue;
                }
                // Pomija, jeśli plik nie jest częścią .ipt.
                if (!sciezkaPliku.ToLower().EndsWith(".ipt")) continue;
                // Pomija, jeśli plik został już przetworzony.
                if (przetworzonePliki.Contains(sciezkaPliku)) continue;
                przetworzonePliki.Add(sciezkaPliku); // Dodaje plik do zbioru przetworzonych.

                string grawer = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
                Image miniatura = null;
                PartDocument partDoc = null;

                // Zamyka otwarte dokumenty, które odpowiadają ścieżce pliku, aby uniknąć problemów z otwieraniem.
                foreach (Document doc in aplikacja.Documents)
                {
                    if (string.Equals(doc.FullFileName, sciezkaPliku, StringComparison.OrdinalIgnoreCase) && doc is PartDocument)
                    {
                        try { doc.Close(false); } catch { }
                        break;
                    }
                }
                // Sprawdza, czy plik fizycznie istnieje.
                if (!System.IO.File.Exists(sciezkaPliku))
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorFileDoesNotExist", sciezkaPliku));
                            break;
                        }
                    }
                    continue;
                }
                try
                {
                    // Otwiera dokument części.
                    partDoc = aplikacja.Documents.Open(sciezkaPliku, true) as PartDocument;
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli otwarcie części się nie powiedzie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorOpeningPart", sciezkaPliku, ex.Message));
                            break;
                        }
                    }
                    continue;
                }
                if (partDoc != null)
                {
                    Document prevActive = aplikacja.ActiveDocument; // Zapisuje aktywny dokument.
                    try
                    {
                        partDoc.Activate(); // Aktywuje dokument części.
                        var camera = aplikacja.ActiveView.Camera; // Pobiera kamerę aktywnego widoku.
                        camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopLeftViewOrientation; // Ustawia orientację widoku.
                        camera.Fit(); // Dopasowuje widok do zawartości.
                        camera.Apply(); // Stosuje zmiany kamery.
                        aplikacja.ActiveView.Update(); // Aktualizuje widok.

                        // Zapisuje zrzut ekranu jako tymczasowy plik BMP.
                        string tempPath = System.IO.Path.GetTempFileName() + ".bmp";
                        aplikacja.ActiveView.SaveAsBitmap(tempPath, 256, 256);
                        // Wczytuje obraz z pliku tymczasowego i konwertuje go na miniaturę.
                        using (var bmp = new Bitmap(tempPath))
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                miniatura = (Image)Image.FromStream(ms).Clone();
                            }
                        }
                        System.IO.File.Delete(tempPath); // Usuwa tymczasowy plik.
                    }
                    catch (Exception ex)
                    {
                        // Loguje błąd, jeśli generowanie miniatury się nie powiedzie.
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage(TranslationManager.GetTranslation("ErrorGeneratingThumbnail", sciezkaPliku, ex.Message));
                                break;
                            }
                        }
                    }
                    finally
                    {
                        // Przywraca poprzednio aktywny dokument i zamyka otwartą część.
                        if (aplikacja.ActiveDocument != prevActive && prevActive != null)
                        {
                            try { prevActive.Activate(); } catch { }
                        }
                        try { partDoc.Close(false); } catch { }
                    }
                }
                // Dodaje informację o części z miniaturą do listy.
                lista.Add(new CzesciInfo { Nazwa = wystapienie.Name, SciezkaPliku = sciezkaPliku, Miniatura = miniatura, Grawer = grawer });
            }
        }

        /// <summary>
        /// Pobiera wartość niestandardowej właściwości użytkownika z dokumentu.
        /// </summary>
        /// <param name="dokument">Dokument Inventor.</param>
        /// <param name="nazwaWlasciwosci">Nazwa właściwości do pobrania.</param>
        /// <returns>Wartość właściwości jako string, lub pusty string, jeśli właściwość nie istnieje lub wystąpi błąd.</returns>
        public static string PobierzWlasciwoscUzytkownika(Document dokument, string nazwaWlasciwosci)
        {
            try
            {
                // Iteruje przez wszystkie zestawy właściwości w dokumencie.
                foreach (PropertySet set in dokument.PropertySets)
                {
                    // Szuka zestawu właściwości użytkownika.
                    if (set.Name == "Inventor User Defined Properties" || set.DisplayName == "Inventor User Defined Properties")
                    {
                        // Iteruje przez właściwości w znalezionym zestawie.
                        foreach (Inventor.Property prop in set)
                        {
                            // Jeśli nazwa właściwości pasuje, zwraca jej wartość.
                            if (prop.Name == nazwaWlasciwosci)
                                return prop.Value?.ToString() ?? "";
                        }
                    }
                }
                return ""; // Zwraca pusty string, jeśli właściwość nie została znaleziona.
            }
            catch
            {
                return ""; // Zwraca pusty string w przypadku błędu.
            }
        }

        /// <summary>
        /// Ustawia wartość niestandardowej właściwości użytkownika w pliku części.
        /// </summary>
        /// <param name="sciezkaPliku">Ścieżka do pliku części.</param>
        /// <param name="nazwaWlasciwosci">Nazwa właściwości do ustawienia.</param>
        /// <param name="nowaWartosc">Nowa wartość właściwości.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        /// <param name="logMessage">Opcjonalna akcja do logowania wiadomości.</param>
        /// <returns>True, jeśli właściwość została ustawiona pomyślnie, false w przeciwnym razie.</returns>
        public static bool UstawWlasciwoscUzytkownika(string sciezkaPliku, string nazwaWlasciwosci, string nowaWartosc, Inventor.Application aplikacja, Action<string> logMessage = null)
        {
            try
            {
                // Otwiera dokument części w trybie tylko do odczytu (false).
                var dokument = aplikacja.Documents.Open(sciezkaPliku, false);
                // Pobiera zestaw właściwości użytkownika.
                var propertySets = dokument.PropertySets["Inventor User Defined Properties"];
                Inventor.Property prop = null;
                try { prop = propertySets[nazwaWlasciwosci] as Inventor.Property; } catch { } // Próbuje pobrać istniejącą właściwość.
                if (prop != null)
                {
                    prop.Value = nowaWartosc; // Ustawia nową wartość, jeśli właściwość istnieje.
                }
                else
                {
                    propertySets.Add(nowaWartosc, nazwaWlasciwosci); // Dodaje nową właściwość, jeśli nie istnieje.
                }
                dokument.Save(); // Zapisuje zmiany w dokumencie.
                dokument.Close(true); // Zamyka dokument, zachowując zmiany.
                return true;
            }
            catch (Exception ex)
            {
                // Loguje błąd, jeśli ustawienie właściwości się nie powiedzie.
                logMessage?.Invoke(TranslationManager.GetTranslation("ErrorSettingUserProperty", sciezkaPliku, ex.Message));
                return false;
            }
        }

        /// <summary>
        /// Proste wykrywanie domyślnej miniatury (np. X): sprawdza, czy obraz jest bardzo mały lub jednolity.
        /// </summary>
        /// <param name="img">Obraz do sprawdzenia.</param>
        /// <returns>True, jeśli obraz jest domyślną miniaturą, false w przeciwnym razie.</returns>
        private static bool IsDefaultThumbnail(Image img)
        {
            if (img == null) return true; // Jeśli obraz jest null, to jest domyślny.
            if (img.Width < 16 || img.Height < 16) return true; // Jeśli obraz jest zbyt mały, to jest domyślny.
            using (var bmp = new Bitmap(img))
            {
                var c = bmp.GetPixel(0, 0); // Pobiera kolor pierwszego piksela.
                // Sprawdza, czy wszystkie piksele mają ten sam kolor (jednolity obraz).
                for (int x = 0; x < bmp.Width; x++)
                    for (int y = 0; y < bmp.Height; y++)
                        if (bmp.GetPixel(x, y) != c)
                            return false; // Jeśli znajdzie inny kolor, nie jest jednolity.
            }
            return true; // Jeśli wszystkie piksele mają ten sam kolor, jest jednolity.
        }
    }

    /// <summary>
    /// Klasa logiki grawerowania, wywoływana z przycisku na formularzu.
    /// </summary>
    public class Grawerowanie
    {
        /// <summary>
        /// Wykonuje logikę grawerowania na aktywnym złożeniu w Inventorze.
        /// </summary>
        /// <param name="fontSize">Rozmiar czcionki do grawerowania.</param>
        /// <param name="zamykajCzesci">Czy zamykać części po grawerowaniu.</param>
        public void DoSomething(float fontSize, bool zamykajCzesci)
        {
            // Pobiera aktywną instancję aplikacji Inventor.
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja == null)
            {
                // Wyświetla komunikat o błędzie, jeśli Inventor nie jest uruchomiony.
                MessageBox.Show(TranslationManager.GetTranslation("ErrorInventorNotRunningMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorInventorNotRunning"));
                        break;
                    }
                }
                return;
            }
            if (aplikacja.ActiveDocument == null)
            {
                // Wyświetla komunikat o błędzie, jeśli żaden dokument nie jest otwarty.
                MessageBox.Show(TranslationManager.GetTranslation("ErrorOpenAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorOpenAssembly"));
                        break;
                    }
                }
                return;
            }
            // Sprawdza, czy aktywny dokument jest złożeniem.
            if (!(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
            {
                // Wyświetla komunikat o błędzie, jeśli aktywny dokument nie jest złożeniem.
                MessageBox.Show(TranslationManager.GetTranslation("ErrorNotAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorNotAssembly"));
                        break;
                    }
                }
                return;
            }
            // Zapamiętaj referencję do złożenia użytego do grawerowania
            var docToReactivate = dokumentZespolu;
            // POBIERZ ZAWSZE Z comboBoxFontSize
            float realFontSize = 1.0f;
            var culture = System.Globalization.CultureInfo.CurrentCulture;
            // Pobiera rozmiar czcionki z kontrolki ComboBox na formularzu głównym.
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mainForm && mainForm.Controls["comboBoxFontSize"] is ComboBox cb && cb.SelectedItem is string selected)
                {
                    // Próbuje sparsować rozmiar czcionki, jeśli się nie uda, ustawia domyślny.
                    if (!float.TryParse(selected, System.Globalization.NumberStyles.Float, culture, out realFontSize) || realFontSize <= 0)
                        realFontSize = 1.0f;
                    mainForm.LogMessage(TranslationManager.GetTranslation("LogFontSizeUsed", realFontSize));
                    break;
                }
            }
            // Sprawdza, czy złożenie zostało zapisane.
            if (string.IsNullOrEmpty(dokumentZespolu.FullFileName))
            {
                // Wyświetla komunikat o błędzie, jeśli złożenie nie zostało zapisane.
                MessageBox.Show(TranslationManager.GetTranslation("ErrorSaveAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorSaveAssembly"));
                        break;
                    }
                }
                return;
            }

            // Pobiera listę części ze złożenia.
            var czesci = CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja);
            // Tworzy ścieżkę do folderu IGES.
            string folderIGES = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dokumentZespolu.FullFileName), "IGES");
            WyczyscFolder(folderIGES, "*.igs"); // Czyści folder IGES.
            var przetworzonePliki = new HashSet<string>();
            var newlyOpenedParts = new List<PartDocument>(); // Lista nowo otwartych części.
            var grawerowaneCzesci = new List<(string Nazwa, string Grawer)>(); // Lista grawerowanych części.

            // Iteruje przez każdą część.
            foreach (var czesc in czesci)
            {
                if (string.IsNullOrEmpty(czesc.Grawer)) continue; // Pomija, jeśli brak tekstu do grawerowania.
                if (przetworzonePliki.Contains(czesc.SciezkaPliku)) continue; // Pomija, jeśli plik już przetworzony.

                bool wasNewlyOpened;
                // Wykonuje grawerowanie na części.
                PartDocument engravedPart = WykonajGrawerowanie(czesc.SciezkaPliku, czesc.Grawer, realFontSize, aplikacja, out wasNewlyOpened);
                if (engravedPart != null && wasNewlyOpened)
                {
                    newlyOpenedParts.Add(engravedPart); // Dodaje nowo otwartą część do listy.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"[GRAWEROWANIE] Dodano do listy do zamknięcia: {engravedPart.FullFileName}");
                            break;
                        }
                    }
                }
                EksportujDoIGES(czesc.SciezkaPliku, folderIGES, aplikacja); // Eksportuje część do formatu IGES.
                grawerowaneCzesci.Add((System.IO.Path.GetFileNameWithoutExtension(czesc.SciezkaPliku), czesc.Grawer)); // Dodaje grawerowaną część do listy.
                przetworzonePliki.Add(czesc.SciezkaPliku); // Dodaje plik do zbioru przetworzonych.
            }

            if (zamykajCzesci) // Przeniesiono logikę zamykania poza pętlę
            {
                // Zamyka nowo otwarte części, jeśli opcja jest włączona.
                foreach (var partDoc in newlyOpenedParts)
                {
                    try
                    {
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage($"[GRAWEROWANIE] Próba zamknięcia części: {partDoc.FullFileName}");
                                break;
                            }
                        }
                        partDoc.Close(false); // Zamyka dokument części bez zapisywania.
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage($"[GRAWEROWANIE] Część zamknięta: {partDoc.FullFileName}");
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Loguje błąd, jeśli zamknięcie części się nie powiedzie.
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage(TranslationManager.GetTranslation("LogErrorClosingPart", partDoc.FullFileName, ex.Message));
                                break;
                            }
                        }
                    }
                }
            }
            if (grawerowaneCzesci.Count > 0)
            {
                // Loguje informacje o grawerowanych częściach.
                var logMsg = TranslationManager.GetTranslation("LogEngravedParts") + System.Environment.NewLine;
                foreach (var (nazwa, grawer) in grawerowaneCzesci)
                {
                    logMsg += $"- {nazwa}: {grawer}" + System.Environment.NewLine;
                }
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(logMsg);
                        break;
                    }
                }
            }
            // Po zakończeniu grawerowania wróć do złożenia użytego do grawerowania
            if (docToReactivate != null)
            {
                try
                {
                    docToReactivate.Activate(); // Aktywuje poprzednio aktywny dokument.
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli powrót do złożenia się nie powiedzie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("LogErrorReturningToAssembly", ex.Message));
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Czyści folder z plików IGES.
        /// </summary>
        /// <param name="sciezkaFolderu">Ścieżka do folderu.</param>
        /// <param name="maskaPliku">Maska plików do usunięcia (np. "*.igs").</param>
        private void WyczyscFolder(string sciezkaFolderu, string maskaPliku)
        {
            if (System.IO.Directory.Exists(sciezkaFolderu))
            {
                // Usuwa wszystkie pliki pasujące do maski w folderze.
                foreach (var plik in System.IO.Directory.GetFiles(sciezkaFolderu, maskaPliku))
                {
                    System.IO.File.Delete(plik);
                }
            }
            else
            {
                System.IO.Directory.CreateDirectory(sciezkaFolderu); // Tworzy folder, jeśli nie istnieje.
            }
        }

        /// <summary>
        /// Wykonuje grawerowanie na pliku części, jeśli posiada właściwość "Grawer".
        /// </summary>
        /// <param name="sciezkaPliku">Ścieżka do pliku części.</param>
        /// <param name="tekstGrawerowania">Tekst do grawerowania.</param>
        /// <param name="fontSize">Rozmiar czcionki.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        /// <param name="wasNewlyOpened">Zwraca true, jeśli dokument został nowo otwarty przez tę metodę.</param>
        /// <returns>Dokument części po grawerowaniu, lub null w przypadku błędu.</returns>
        public PartDocument WykonajGrawerowanie(string sciezkaPliku, string tekstGrawerowania, float fontSize, Inventor.Application aplikacja, out bool wasNewlyOpened)
        {
            wasNewlyOpened = false;
            if (string.IsNullOrEmpty(tekstGrawerowania)) return null; // Pomija, jeśli brak tekstu do grawerowania.

            PartDocument dokumentCzesci = null;
            // Sprawdza, czy dokument jest już otwarty.
            foreach (Document doc in aplikacja.Documents)
            {
                if (string.Equals(doc.FullFileName, sciezkaPliku, StringComparison.OrdinalIgnoreCase) && doc is PartDocument)
                {
                    dokumentCzesci = (PartDocument)doc;
                    break;
                }
            }

            if (dokumentCzesci == null)
            {
                // Dokument nie jest otwarty, otwórz go.
                try
                {
                    dokumentCzesci = aplikacja.Documents.Open(sciezkaPliku, true) as PartDocument; // Otwiera widocznie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"[GRAWEROWANIE] Otwarto część widocznie: {sciezkaPliku}");
                            break;
                        }
                    }
                    wasNewlyOpened = true; // Ustawia flagę, że dokument został nowo otwarty.
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli otwarcie części do grawerowania się nie powiedzie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorOpeningPartForEngraving", sciezkaPliku, ex.Message));
                            break;
                        }
                    }
                    return null;
                }
            }

            if (dokumentCzesci == null) return null;

            // Ustawia widok kamery do grawerowania.
            Document prevActive = aplikacja.ActiveDocument; // Zapisuje aktywny dokument.
            try
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Aktywowanie dokumentu: {dokumentCzesci.FullFileName}");
                        break;
                    }
                }
                dokumentCzesci.Activate(); // Aktywuje dokument części.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Dokument aktywowany. Próba ustawienia kamery.");
                        break;
                    }
                }
                var camera = aplikacja.ActiveView.Camera; // Pobiera kamerę aktywnego widoku.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Ustawianie ViewOrientationType na kIsoTopLeftViewOrientation.");
                        break;
                    }
                }
                camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopLeftViewOrientation; // Ustawia orientację widoku.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Wywoływanie Fit().");
                        break;
                    }
                }
                camera.Fit(); // Dopasowuje widok do zawartości.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Wywoływanie Apply().");
                        break;
                    }
                }
                camera.Apply(); // Stosuje zmiany kamery.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Wywoływanie aplikacja.ActiveView.Update().");
                        break;
                    }
                }
                aplikacja.ActiveView.Update(); // Aktualizuje widok.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Kamera ustawiona pomyślnie.");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Loguje błąd, jeśli ustawienie kamery się nie powiedzie.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorSettingCamera", dokumentCzesci.FullFileName, ex.Message));
                        break;
                    }
                }
            }
            finally
            {
                // Przywraca poprzednio aktywny dokument.
                if (aplikacja.ActiveDocument != prevActive && prevActive != null)
                {
                    try { prevActive.Activate(); } catch { }
                }
            }

            if (dokumentCzesci.ComponentDefinition == null)
            {
                // Loguje błąd, jeśli definicja komponentu części jest brakująca.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorPartNoComponentDefinition", sciezkaPliku));
                        break;
                    }
                }
                return null;
            }
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] ComponentDefinition dla {sciezkaPliku}: {dokumentCzesci.ComponentDefinition != null}");
                    break;
                }
            }

            if (dokumentCzesci.ComponentDefinition.WorkPlanes.Count < 3)
            {
                // Loguje błąd, jeśli część ma mniej niż 3 płaszczyzny robocze.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorPartFewerThan3WorkPlanes", sciezkaPliku));
                        break;
                    }
                }
                return null;
            }
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] WorkPlanes.Count dla {sciezkaPliku}: {dokumentCzesci.ComponentDefinition.WorkPlanes.Count}");
                    break;
                }
            }

            PlanarSketch szkic = null;
            try
            {
                // Dodaje nowy szkic na drugiej płaszczyźnie roboczej.
                szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[2]);
            }
            catch (Exception ex)
            {
                // Loguje błąd, jeśli nie można utworzyć szkicu.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateSketch", 2, sciezkaPliku) + $" - {ex.Message}");
                        break;
                    }
                }
                return null;
            }
            if (szkic == null)
            {
                // Loguje błąd, jeśli szkic jest null.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateSketch", 2, sciezkaPliku));
                        break;
                    }
                }
                return null;
            }
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] Szkic utworzony dla {sciezkaPliku}: {szkic != null}");
                    break;
                }
            }

            Point2d pozycjaTekstu = aplikacja.TransientGeometry.CreatePoint2d(-1, 0); // Tworzy punkt dla pozycji tekstu.
            string fontName = "Tahoma"; // Domyślna nazwa czcionki.
            float realFontSize = fontSize; // Rzeczywisty rozmiar czcionki.
            // Pobiera nazwę czcionki z kontrolki ComboBox na formularzu głównym.
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mainForm)
                {
                    if (mainForm.Controls["comboBoxFonts"] is ComboBox cb && cb.SelectedItem is string selectedFont)
                        fontName = selectedFont;
                    break;
                }
            }
            var currentCulture = System.Globalization.CultureInfo.CurrentCulture;
            // Tworzy sformatowany tekst do grawerowania.
            string formattedText = "<StyleOverride Font='" + fontName + "' FontSize='" + realFontSize.ToString("0.0", currentCulture) + "'>" + tekstGrawerowania + "</StyleOverride>";
            string debugText = formattedText + " | Path: " + sciezkaPliku;
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage(TranslationManager.GetTranslation("DebugStyleOverride", debugText));
                    break;
                }
            }
            Inventor.TextBox poleTekstowe = null;
            try
            {
                // Dodaje pole tekstowe do szkicu.
                poleTekstowe = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
            }
            catch (Exception ex)
            {
                // Loguje błąd, jeśli nie można utworzyć pola tekstowego.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateTextBox", sciezkaPliku) + $" - {ex.Message}");
                        break;
                    }
                }
                return null;
            }
            if (poleTekstowe == null)
            {
                // Loguje błąd, jeśli pole tekstowe jest null.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateTextBox", sciezkaPliku));
                        break;
                    }
                }
                return null;
            }
            poleTekstowe.FormattedText = formattedText; // Ustawia sformatowany tekst.
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] Pole tekstowe utworzone dla {sciezkaPliku}: {poleTekstowe != null}");
                    break;
                }
            }

            var funkcjeEkstruzji = dokumentCzesci.ComponentDefinition.Features.ExtrudeFeatures;
            // Sprawdza, czy istnieje już operacja ekstruzji o nazwie "Grawerowanie64" i usuwa ją.
            foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
            {
                if (operacjaEkstrudowania.Name == "Grawerowanie64")
                {
                    operacjaEkstrudowania.Delete();
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"[GRAWEROWANIE] Usunięto istniejące grawerowanie 'Grawerowanie64' dla {sciezkaPliku}.");
                            break;
                        }
                    }
                    break;
                }
            }

            Profile profil = null;
            try
            {
                // Tworzy profil ze szkicu.
                profil = szkic.Profiles.AddForSolid();
            }
            catch (Exception ex)
            {
                // Loguje błąd, jeśli nie można utworzyć profilu.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateProfile", sciezkaPliku) + $" - {ex.Message}");
                        break;
                    }
                }
                return null;
            }
            if (profil == null)
            {
                // Loguje błąd, jeśli profil jest null.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateProfile", sciezkaPliku));
                        break;
                    }
                }
                return null;
            }
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] Profil utworzony dla {sciezkaPliku}: {profil != null}");
                    break;
                }
            }

            ExtrudeFeature operacjaEkstrudowaniaNowa = null;
            try
            {
                // Tworzy nową operację ekstruzji (wycięcia) na podstawie profilu.
                operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                    profil,
                    PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                    PartFeatureOperationEnum.kCutOperation
                );
            }
            catch (Exception ex)
            {
                // Loguje błąd, jeśli nie można utworzyć ekstruzji.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateExtrusion", sciezkaPliku) + $" - {ex.Message}");
                        break;
                    }
                }
                return null;
            }

            if (operacjaEkstrudowaniaNowa == null)
            {
                // Loguje błąd, jeśli operacja ekstruzji jest null.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateExtrusion", sciezkaPliku));
                        break;
                    }
                }
                return null;
            }
            operacjaEkstrudowaniaNowa.Name = "Grawerowanie64"; // Nadaje nazwę operacji ekstruzji.
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] Operacja ekstruzji utworzona dla {sciezkaPliku}: {operacjaEkstrudowaniaNowa != null}");
                    break;
                }
            }

            // Sprawdza status zdrowia operacji ekstruzji (może wskazywać na błąd płaszczyzny szkicu).
            if ((int)operacjaEkstrudowaniaNowa.HealthStatus == 11780)
            {
                // Loguje informację o błędnej płaszczyźnie szkicu.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("LogWrongSketchPlane"));
                        break;
                    }
                }
                // Usuwa poprzednią operację grawerowania.
                foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
                {
                    if (operacjaEkstrudowania.Name == "Grawerowanie64")
                    {
                        operacjaEkstrudowania.Delete();
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage($"[GRAWEROWANIE] Usunięto istniejące grawerowanie 'Grawerowanie64' dla {sciezkaPliku} (druga próba).");
                                break;
                            }
                        }
                        break;
                    }
                }
                szkic = null;
                try
                {
                    // Próbuje utworzyć szkic na innej płaszczyźnie roboczej (trzeciej).
                    szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[3]);
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli nie można utworzyć szkicu na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateSketch", 3, sciezkaPliku) + $" - {ex.Message}");
                            break;
                        }
                    }
                    return null;
                }
                if (szkic == null)
                {
                    // Loguje błąd, jeśli szkic jest null na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateSketch", 3, sciezkaPliku));
                            break;
                        }
                    }
                    return null;
                }
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Szkic utworzony dla {sciezkaPliku} (druga próba): {szkic != null}");
                        break;
                    }
                }

                Inventor.TextBox poleTekstowe2 = null;
                try
                {
                    // Dodaje pole tekstowe do nowego szkicu.
                    poleTekstowe2 = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli nie można utworzyć pola tekstowego na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateTextBox", sciezkaPliku) + $" on second attempt. - {ex.Message}");
                            break;
                        }
                    }
                    return null;
                }
                if (poleTekstowe2 == null)
                {
                    // Loguje błąd, jeśli pole tekstowe jest null na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateTextBox", sciezkaPliku) + " on second attempt.");
                            break;
                        }
                    }
                    return null;
                }

                string formattedText2 = "<StyleOverride Font='" + fontName + "' FontSize='" + realFontSize.ToString("0.0", currentCulture) + "'>" + tekstGrawerowania + "</StyleOverride>";
                string debugText2 = formattedText2 + " | Path: " + sciezkaPliku;
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("DebugStyleOverride", debugText2));
                        break;
                    }
                }
                poleTekstowe2.FormattedText = formattedText2; // Ustawia sformatowany tekst.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Pole tekstowe utworzone dla {sciezkaPliku} (druga próba): {poleTekstowe2 != null}");
                        break;
                    }
                }

                Profile nowyProfil = null;
                try
                {
                    // Tworzy nowy profil ze szkicu.
                    nowyProfil = szkic.Profiles.AddForSolid();
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli nie można utworzyć profilu na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateProfile", sciezkaPliku) + $" on second attempt. - {ex.Message}");
                            break;
                        }
                    }
                    return null;
                }
                if (nowyProfil == null)
                {
                    // Loguje błąd, jeśli nowy profil jest null na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateProfile", sciezkaPliku) + " on second attempt.");
                            break;
                        }
                    }
                    return null;
                }
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Profil utworzony dla {sciezkaPliku} (druga próba): {nowyProfil != null}");
                        break;
                    }
                }

                operacjaEkstrudowaniaNowa = null;
                try
                {
                    // Tworzy nową operację ekstruzji (wycięcia) na podstawie nowego profilu.
                    operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                        nowyProfil,
                        PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                        PartFeatureOperationEnum.kCutOperation
                    );
                }
                catch (Exception ex)
                {
                    // Loguje błąd, jeśli nie można utworzyć ekstruzji na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateExtrusion", sciezkaPliku) + $" on second attempt. - {ex.Message}");
                            break;
                        }
                    }
                    return null;
                }

                if (operacjaEkstrudowaniaNowa == null)
                {
                    // Loguje błąd, jeśli operacja ekstruzji jest null na drugiej próbie.
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorCouldNotCreateExtrusion", sciezkaPliku) + " on second attempt.");
                            break;
                        }
                    }
                    return null;
                }
                operacjaEkstrudowaniaNowa.Name = "Grawerowanie64"; // Nadaje nazwę operacji ekstruzji.
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Operacja ekstruzji utworzona dla {sciezkaPliku} (druga próba): {operacjaEkstrudowaniaNowa != null}");
                        break;
                    }
                }
            }

            dokumentCzesci.Save(); // Zapisuje zmiany w dokumencie części.
            return dokumentCzesci; // Zwraca dokument części.
        }

        /// <summary>
        /// Eksportuje plik części do formatu IGES do określonego folderu.
        /// </summary>
        /// <param name="sciezkaPliku">Ścieżka do pliku części.</param>
        /// <param name="folderIGES">Ścieżka do folderu docelowego IGES.</param>
        /// <param name="aplikacja">Instancja aplikacji Inventor.</param>
        private void EksportujDoIGES(string sciezkaPliku, string folderIGES, Inventor.Application aplikacja)
        {
            if (!System.IO.File.Exists(sciezkaPliku)) return; // Pomija, jeśli plik nie istnieje.
            var dokument = aplikacja.Documents.Open(sciezkaPliku); // Otwiera dokument.
            if (dokument == null) return; // Pomija, jeśli dokument jest null.
            string tekstGrawerowania = CzesciHelper.PobierzWlasciwoscUzytkownika(dokument, "Grawer"); // Pobiera tekst grawerowania.
            if (string.IsNullOrEmpty(tekstGrawerowania)) return; // Pomija, jeśli brak tekstu do grawerowania.
            // Pobiera dodatek tłumacza IGES.
            var tlumaczIGES = aplikacja.ApplicationAddIns.ItemById["{90AF7F44-0C01-11D5-8E83-0010B541CD80}"] as TranslatorAddIn;
            if (tlumaczIGES == null) return; // Pomija, jeśli tłumacz IGES nie jest dostępny.
            var kontekstTlumaczenia = aplikacja.TransientObjects.CreateTranslationContext(); // Tworzy kontekst tłumaczenia.
            kontekstTlumaczenia.Type = IOMechanismEnum.kFileBrowseIOMechanism; // Ustawia typ mechanizmu I/O.
            // ...tutaj można dodać logikę eksportu IGES...
        }
    }
}
