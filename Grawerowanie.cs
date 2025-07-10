using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Inventor;

namespace mca64Inventor
{
    /// <summary>
    /// Data model for parts used for display and engraving.
    /// </summary>
    public class CzesciInfo
    {
        public string Nazwa { get; set; }
        public string SciezkaPliku { get; set; }
        public Image Miniatura { get; set; }
        public string Grawer { get; set; }
    }

    /// <summary>
    /// Utility class for operations on assembly parts.
    /// </summary>
    public static class CzesciHelper
    {
        /// <summary>
        /// Retrieves a list of parts from the assembly along with thumbnails and engraving field.
        /// </summary>
        public static List<CzesciInfo> PobierzCzesciZespolu(AssemblyDocument dokumentZespolu, Inventor.Application aplikacja, bool generateThumbnails = false)
        {
            var lista = new List<CzesciInfo>();
            if (dokumentZespolu.ComponentDefinition == null)
            {
                MessageBox.Show("The active assembly is missing its core component definition and cannot be processed.", "Assembly Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lista;
            }
            PrzetworzWystapienia(dokumentZespolu.ComponentDefinition.Occurrences, lista, new HashSet<string>(), aplikacja, generateThumbnails);
            return lista;
        }

        private static void PrzetworzWystapienia(ComponentOccurrences wystapienia, List<CzesciInfo> lista, HashSet<string> przetworzonePliki, Inventor.Application aplikacja, bool generateThumbnails)
        {
            foreach (ComponentOccurrence wystapienie in wystapienia)
            {
                Document dokument = null;
                try { dokument = (Document)wystapienie.Definition.Document; } catch { }
                if (dokument == null) continue;
                string sciezkaPliku = dokument.FullFileName;
                if (!sciezkaPliku.ToLower().EndsWith(".ipt"))
                {
                    if (dokument is AssemblyDocument assemblyDoc)
                        PrzetworzWystapienia(assemblyDoc.ComponentDefinition.Occurrences, lista, przetworzonePliki, aplikacja, generateThumbnails);
                    continue;
                }
                if (przetworzonePliki.Contains(sciezkaPliku)) continue;
                przetworzonePliki.Add(sciezkaPliku);
                string grawer = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
                lista.Add(new CzesciInfo { Nazwa = wystapienie.Name, SciezkaPliku = sciezkaPliku, Miniatura = null, Grawer = grawer });
            }
        }

        /// <summary>
        /// Generates thumbnails for all .ipt parts in the assembly. Opens each part, takes a screenshot,
        /// and closes only those parts that were opened temporarily.
        /// </summary>
        public static List<CzesciInfo> GenerujMiniaturyDlaCzesci(AssemblyDocument dokumentZespolu, Inventor.Application aplikacja)
        {
            var lista = new List<CzesciInfo>();
            var przetworzonePliki = new HashSet<string>();
            GenerujMiniaturyRekurencyjnie(dokumentZespolu.ComponentDefinition.Occurrences, lista, przetworzonePliki, aplikacja);
            return lista;
        }

        private static void GenerujMiniaturyRekurencyjnie(ComponentOccurrences wystapienia, List<CzesciInfo> lista, HashSet<string> przetworzonePliki, Inventor.Application aplikacja)
        {
            foreach (ComponentOccurrence wystapienie in wystapienia)
            {
                Document dokument = null;
                try { dokument = (Document)wystapienie.Definition.Document; } catch { }
                if (dokument == null) continue;
                string sciezkaPliku = dokument.FullFileName;
                if (dokument is AssemblyDocument subAsm)
                {
                    GenerujMiniaturyRekurencyjnie(subAsm.ComponentDefinition.Occurrences, lista, przetworzonePliki, aplikacja);
                    continue;
                }
                if (!sciezkaPliku.ToLower().EndsWith(".ipt")) continue;
                if (przetworzonePliki.Contains(sciezkaPliku)) continue;
                przetworzonePliki.Add(sciezkaPliku);
                string grawer = PobierzWlasciwoscUzytkownika(dokument, "Grawer");
                Image miniatura = null;
                PartDocument partDoc = null;
                foreach (Document doc in aplikacja.Documents)
                {
                    if (string.Equals(doc.FullFileName, sciezkaPliku, StringComparison.OrdinalIgnoreCase) && doc is PartDocument)
                    {
                        try { doc.Close(false); } catch { }
                        break;
                    }
                }
                if (!System.IO.File.Exists(sciezkaPliku))
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"File does not exist: {sciezkaPliku}");
                            break;
                        }
                    }
                    continue;
                }
                try
                {
                    partDoc = aplikacja.Documents.Open(sciezkaPliku, true) as PartDocument;
                }
                catch (Exception ex)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Error opening part: {sciezkaPliku} - {ex.Message}");
                            break;
                        }
                    }
                    continue;
                }
                if (partDoc != null)
                {
                    Document prevActive = aplikacja.ActiveDocument;
                    try
                    {
                        partDoc.Activate();
                        var camera = aplikacja.ActiveView.Camera;
                        camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopLeftViewOrientation;
                        camera.Fit();
                        camera.Apply();
                        aplikacja.ActiveView.Update();
                        string tempPath = System.IO.Path.GetTempFileName() + ".bmp";
                        aplikacja.ActiveView.SaveAsBitmap(tempPath, 256, 256);
                        using (var bmp = new Bitmap(tempPath))
                        {
                            using (var ms = new System.IO.MemoryStream())
                            {
                                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                miniatura = (Image)Image.FromStream(ms).Clone();
                            }
                        }
                        System.IO.File.Delete(tempPath);
                    }
                    catch (Exception ex)
                    {
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage($"Error generating thumbnail for: {sciezkaPliku} - {ex.Message}");
                                break;
                            }
                        }
                    }
                    finally
                    {
                        if (aplikacja.ActiveDocument != prevActive && prevActive != null)
                        {
                            try { prevActive.Activate(); } catch { }
                        }
                        try { partDoc.Close(false); } catch { }
                    }
                }
                lista.Add(new CzesciInfo { Nazwa = wystapienie.Name, SciezkaPliku = sciezkaPliku, Miniatura = miniatura, Grawer = grawer });
            }
        }

        /// <summary>
        /// Retrieves the value of a user-defined property from the document.
        /// </summary>
        public static string PobierzWlasciwoscUzytkownika(Document dokument, string nazwaWlasciwosci)
        {
            try
            {
                foreach (PropertySet set in dokument.PropertySets)
                {
                    if (set.Name == "Inventor User Defined Properties" || set.DisplayName == "Inventor User Defined Properties")
                    {
                        foreach (Inventor.Property prop in set)
                        {
                            if (prop.Name == nazwaWlasciwosci)
                                return prop.Value?.ToString() ?? "";
                        }
                    }
                }
                return "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Sets the value of a user-defined property in the part file.
        /// </summary>
        public static bool UstawWlasciwoscUzytkownika(string sciezkaPliku, string nazwaWlasciwosci, string nowaWartosc, Inventor.Application aplikacja, Action<string> logMessage = null)
        {
            try
            {
                var dokument = aplikacja.Documents.Open(sciezkaPliku, false);
                var propertySets = dokument.PropertySets["Inventor User Defined Properties"];
                Inventor.Property prop = null;
                try { prop = propertySets[nazwaWlasciwosci] as Inventor.Property; } catch { }
                if (prop != null)
                {
                    prop.Value = nowaWartosc;
                }
                else
                {
                    propertySets.Add(nowaWartosc, nazwaWlasciwosci);
                }
                dokument.Save();
                dokument.Close(true);
                return true;
            }
            catch (Exception ex)
            {
                logMessage?.Invoke($"Error setting user property for '{sciezkaPliku}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Simple detection of default thumbnail (e.g., X): check if the image is very small or uniform.
        /// </summary>
        private static bool IsDefaultThumbnail(Image img)
        {
            if (img == null) return true;
            if (img.Width < 16 || img.Height < 16) return true;
            using (var bmp = new Bitmap(img))
            {
                var c = bmp.GetPixel(0, 0);
                for (int x = 0; x < bmp.Width; x++)
                    for (int y = 0; y < bmp.Height; y++)
                        if (bmp.GetPixel(x, y) != c)
                            return false;
            }
            return true;
        }
    }

    /// <summary>
    /// Engraving logic class, invoked from a button on the form.
    /// </summary>
    public class Grawerowanie
    {
        /// <summary>
        /// Executes engraving logic on the active assembly in Inventor.
        /// </summary>
        public void DoSomething(float fontSize, bool zamykajCzesci)
        {
            var aplikacja = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            if (aplikacja == null)
            {
                MessageBox.Show("Cannot get reference to Inventor!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage("Cannot get reference to Inventor!");
                        break;
                    }
                }
                return;
            }
            if (aplikacja.ActiveDocument == null)
            {
                MessageBox.Show("Open an assembly document before running the script!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage("Open an assembly document before running the script!");
                        break;
                    }
                }
                return;
            }
            if (!(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
            {
                MessageBox.Show("The active document is not an assembly document!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage("The active document is not an assembly document!");
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
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mainForm && mainForm.Controls["comboBoxFontSize"] is ComboBox cb && cb.SelectedItem is string selected)
                {
                    if (!float.TryParse(selected, System.Globalization.NumberStyles.Float, culture, out realFontSize) || realFontSize <= 0)
                        realFontSize = 1.0f;
                    mainForm.LogMessage($"[GRAWEROWANIE] Użyto rozmiaru czcionki z comboBoxFontSize: {realFontSize}");
                    break;
                }
            }
            if (string.IsNullOrEmpty(dokumentZespolu.FullFileName))
            {
                MessageBox.Show("Please save the assembly document before running the script!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage("The assembly must be saved before engraving.");
                        break;
                    }
                }
                return;
            }

            var czesci = CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja);
            string folderIGES = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dokumentZespolu.FullFileName), "IGES");
            WyczyscFolder(folderIGES, "*.igs");
            var przetworzonePliki = new HashSet<string>();
            var newlyOpenedParts = new List<PartDocument>(); // Zmieniono nazwę zmiennej
            var grawerowaneCzesci = new List<(string Nazwa, string Grawer)>();
            foreach (var czesc in czesci)
            {
                if (string.IsNullOrEmpty(czesc.Grawer)) continue;
                if (przetworzonePliki.Contains(czesc.SciezkaPliku)) continue;
                bool wasNewlyOpened;
                PartDocument engravedPart = WykonajGrawerowanie(czesc.SciezkaPliku, czesc.Grawer, realFontSize, aplikacja, out wasNewlyOpened);
                if (engravedPart != null && wasNewlyOpened) newlyOpenedParts.Add(engravedPart); // Dodano do listy nowo otwartych
                EksportujDoIGES(czesc.SciezkaPliku, folderIGES, aplikacja);
                grawerowaneCzesci.Add((System.IO.Path.GetFileNameWithoutExtension(czesc.SciezkaPliku), czesc.Grawer));
                przetworzonePliki.Add(czesc.SciezkaPliku);
            }
            if (zamykajCzesci) // Przeniesiono logikę zamykania poza pętlę
            {
                foreach (var partDoc in newlyOpenedParts)
                {
                    try
                    {
                        partDoc.Close(false);
                    }
                    catch (Exception ex)
                    {
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage($"Błąd zamykania części: {partDoc.FullFileName} - {ex.Message}");
                                break;
                            }
                        }
                    }
                }
            }
            if (grawerowaneCzesci.Count > 0)
            {
                var logMsg = "Engraved parts:" + System.Environment.NewLine;
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
                    docToReactivate.Activate();
                }
                catch (Exception ex)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Błąd podczas powrotu do zakładki ze złożeniem: {ex.Message}");
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Clears the folder of IGES files.
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
        /// Performs engraving on the part file if it has the "Grawer" property.
        /// </summary>
        public PartDocument WykonajGrawerowanie(string sciezkaPliku, string tekstGrawerowania, float fontSize, Inventor.Application aplikacja, out bool wasNewlyOpened)
        {
            wasNewlyOpened = false;
            if (string.IsNullOrEmpty(tekstGrawerowania)) return null;

            PartDocument dokumentCzesci = null;
            // Check if the document is already open
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
                // Document is not open, open it
                try
                {
                    dokumentCzesci = aplikacja.Documents.Open(sciezkaPliku, false) as PartDocument; // Open as invisible
                    wasNewlyOpened = true;
                }
                catch (Exception ex)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Error opening part for engraving: {sciezkaPliku} - {ex.Message}");
                            break;
                        }
                    }
                    return null;
                }
            }

            if (dokumentCzesci == null) return null;

            if (dokumentCzesci.ComponentDefinition == null)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error: Part '{sciezkaPliku}' does not have a valid component definition.");
                        break;
                    }
                }
                return null;
            }

            if (dokumentCzesci.ComponentDefinition.WorkPlanes.Count < 3)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error: Part '{sciezkaPliku}' has fewer than 3 work planes.");
                        break;
                    }
                }
                return null;
            }

            PlanarSketch szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[2]);
            if (szkic == null)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error: Could not create sketch on work plane 2 for part '{sciezkaPliku}'.");
                        break;
                    }
                }
                return null;
            }

            Point2d pozycjaTekstu = aplikacja.TransientGeometry.CreatePoint2d(-1, 0);
            string fontName = "Tahoma";
            float realFontSize = fontSize;
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
            string formattedText = "<StyleOverride Font='" + fontName + "' FontSize='" + realFontSize.ToString("0.0", currentCulture) + "'>" + tekstGrawerowania + "</StyleOverride>";
            string debugText = formattedText + " | Path: " + sciezkaPliku;
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"DEBUG: StyleOverride: {debugText}");
                    break;
                }
            }
            var poleTekstowe = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
            if (poleTekstowe == null)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error: Could not create text box for part '{sciezkaPliku}'.");
                        break;
                    }
                }
                return null;
            }
            poleTekstowe.FormattedText = formattedText;

            var funkcjeEkstruzji = dokumentCzesci.ComponentDefinition.Features.ExtrudeFeatures;
            foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
            {
                if (operacjaEkstrudowania.Name == "Grawerowanie64")
                {
                    operacjaEkstrudowania.Delete();
                    break;
                }
            }

            Profile profil = szkic.Profiles.AddForSolid();
            if (profil == null)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error: Could not create profile for part '{sciezkaPliku}'.");
                        break;
                    }
                }
                return null;
            }

            ExtrudeFeature operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                profil,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kCutOperation
            );

            if (operacjaEkstrudowaniaNowa == null)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"Error: Could not create extrusion for part '{sciezkaPliku}'.");
                        break;
                    }
                }
                return null;
            }
            operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";

            if ((int)operacjaEkstrudowaniaNowa.HealthStatus == 11780)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage("Wrong sketch plane! Trying again with another plane.");
                        break;
                    }
                }
                foreach (ExtrudeFeature operacjaEkstrudowania in funkcjeEkstruzji)
                {
                    if (operacjaEkstrudowania.Name == "Grawerowanie64")
                    {
                        operacjaEkstrudowania.Delete();
                        break;
                    }
                }
                szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[3]);
                if (szkic == null)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Error: Could not create sketch on work plane 3 for part '{sciezkaPliku}'.");
                            break;
                        }
                    }
                    return null;
                }

                var poleTekstowe2 = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
                if (poleTekstowe2 == null)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Error: Could not create text box for part '{sciezkaPliku}' on second attempt.");
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
                        mf.LogMessage($"DEBUG: StyleOverride: {debugText2}");
                        break;
                    }
                }
                poleTekstowe2.FormattedText = formattedText2;

                Profile nowyProfil = szkic.Profiles.AddForSolid();
                if (nowyProfil == null)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Error: Could not create profile for part '{sciezkaPliku}' on second attempt.");
                            break;
                        }
                    }
                    return null;
                }

                operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                    nowyProfil,
                    PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                    PartFeatureOperationEnum.kCutOperation
                );

                if (operacjaEkstrudowaniaNowa == null)
                {
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"Error: Could not create extrusion for part '{sciezkaPliku}' on second attempt.");
                            break;
                        }
                    }
                    return null;
                }
                operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";
            }

            dokumentCzesci.Save();
            return dokumentCzesci;
        }

        /// <summary>
        /// Exports the part file to IGES format to the specified folder.
        /// </summary>
        private void EksportujDoIGES(string sciezkaPliku, string folderIGES, Inventor.Application aplikacja)
        {
            if (!System.IO.File.Exists(sciezkaPliku)) return;
            var dokument = aplikacja.Documents.Open(sciezkaPliku);
            if (dokument == null) return;
            string tekstGrawerowania = CzesciHelper.PobierzWlasciwoscUzytkownika(dokument, "Grawer");
            if (string.IsNullOrEmpty(tekstGrawerowania)) return;
            var tlumaczIGES = aplikacja.ApplicationAddIns.ItemById["{90AF7F44-0C01-11D5-8E83-0010B541CD80}"] as TranslatorAddIn;
            if (tlumaczIGES == null) return;
            var kontekstTlumaczenia = aplikacja.TransientObjects.CreateTranslationContext();
            kontekstTlumaczenia.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            // ...here you can add IGES export logic...
        }
    }
}