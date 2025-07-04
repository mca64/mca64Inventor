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
                            miniatura = new Bitmap(bmp);
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
        public static bool UstawWlasciwoscUzytkownika(string sciezkaPliku, string nazwaWlasciwosci, string nowaWartosc, Inventor.Application aplikacja)
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
            catch
            {
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
            // POBIERZ ZAWSZE Z comboBoxFontSize
            float realFontSize = 1.0f;
            var culture = System.Globalization.CultureInfo.CurrentCulture;
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mainForm && mainForm.Controls["comboBoxFontSize"] is ComboBox cb && cb.SelectedItem is string selected)
                {
                    if (!float.TryParse(selected, System.Globalization.NumberStyles.Float, culture, out realFontSize) || realFontSize <= 0)
                        realFontSize = 1.0f;
                    mainForm.LogMessage($"[GRAWEROWANIE] U¿yto rozmiaru czcionki z comboBoxFontSize: {realFontSize}");
                    break;
                }
            }
            var czesci = CzesciHelper.PobierzCzesciZespolu(dokumentZespolu, aplikacja);
            string folderIGES = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dokumentZespolu.FullFileName), "IGES");
            WyczyscFolder(folderIGES, "*.igs");
            var przetworzonePliki = new HashSet<string>();
            var otwarteDokumenty = new List<Document>();
            var grawerowaneCzesci = new List<(string Nazwa, string Grawer)>();
            foreach (var czesc in czesci)
            {
                if (string.IsNullOrEmpty(czesc.Grawer)) continue;
                if (przetworzonePliki.Contains(czesc.SciezkaPliku)) continue;
                WykonajGrawerowanie(czesc.SciezkaPliku, czesc.Grawer, realFontSize, aplikacja);
                EksportujDoIGES(czesc.SciezkaPliku, folderIGES, aplikacja);
                if (zamykajCzesci)
                {
                    foreach (Document doc in aplikacja.Documents)
                    {
                        if (string.Equals(doc.FullFileName, czesc.SciezkaPliku, StringComparison.OrdinalIgnoreCase))
                        {
                            try { doc.Close(false); } catch { }
                            break;
                        }
                    }
                }
                grawerowaneCzesci.Add((System.IO.Path.GetFileNameWithoutExtension(czesc.SciezkaPliku), czesc.Grawer));
                przetworzonePliki.Add(czesc.SciezkaPliku);
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
        public void WykonajGrawerowanie(string sciezkaPliku, string tekstGrawerowania, float fontSize, Inventor.Application aplikacja)
        {
            if (string.IsNullOrEmpty(tekstGrawerowania)) return;
            var dokumentCzesci = aplikacja.Documents.Open(sciezkaPliku) as PartDocument;
            if (dokumentCzesci == null) return;

            PlanarSketch szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[2]);
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
            ExtrudeFeature operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                profil,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kCutOperation
            );
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
                var poleTekstowe2 = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
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
                operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                    nowyProfil,
                    PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                    PartFeatureOperationEnum.kCutOperation
                );
                operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";
            }

            dokumentCzesci.Save();
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
