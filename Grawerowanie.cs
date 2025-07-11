using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Inventor;
using System.Globalization; // Added for CultureInfo

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
                MessageBox.Show(TranslationManager.GetTranslation("ErrorAssemblyMissingComponentDefinition"), TranslationManager.GetTranslation("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorFileDoesNotExist", sciezkaPliku));
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
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorOpeningPart", sciezkaPliku, ex.Message));
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
                                mf.LogMessage(TranslationManager.GetTranslation("ErrorGeneratingThumbnail", sciezkaPliku, ex.Message));
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
                logMessage?.Invoke(TranslationManager.GetTranslation("ErrorSettingUserProperty", sciezkaPliku, ex.Message));
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
            if (!(aplikacja.ActiveDocument is AssemblyDocument dokumentZespolu))
            {
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
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mainForm && mainForm.Controls["comboBoxFontSize"] is ComboBox cb && cb.SelectedItem is string selected)
                {
                    if (!float.TryParse(selected, System.Globalization.NumberStyles.Float, culture, out realFontSize) || realFontSize <= 0)
                        realFontSize = 1.0f;
                    mainForm.LogMessage(TranslationManager.GetTranslation("LogFontSizeUsed", realFontSize));
                    break;
                }
            }
            if (string.IsNullOrEmpty(dokumentZespolu.FullFileName))
            {
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
                if (engravedPart != null && wasNewlyOpened)
                {
                    newlyOpenedParts.Add(engravedPart);
                    foreach (Form f in System.Windows.Forms.Application.OpenForms)
                    {
                        if (f is MainForm mf)
                        {
                            mf.LogMessage($"[GRAWEROWANIE] Dodano do listy do zamknięcia: {engravedPart.FullFileName}");
                            break;
                        }
                    }
                } // Dodano do listy nowo otwartych
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
                        foreach (Form f in System.Windows.Forms.Application.OpenForms)
                        {
                            if (f is MainForm mf)
                            {
                                mf.LogMessage($"[GRAWEROWANIE] Próba zamknięcia części: {partDoc.FullFileName}");
                                break;
                            }
                        }
                        partDoc.Close(false);
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
                    docToReactivate.Activate();
                }
                catch (Exception ex)
                {
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
                            mf.LogMessage(TranslationManager.GetTranslation("ErrorOpeningPartForEngraving", sciezkaPliku, ex.Message));
                            break;
                        }
                    }
                    return null;
                }
            }

            if (dokumentCzesci == null) return null;

            // Set camera view for engraving
            Document prevActive = aplikacja.ActiveDocument;
            try
            {
                dokumentCzesci.Activate();
                var camera = aplikacja.ActiveView.Camera;
                camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopLeftViewOrientation;
                camera.Fit();
                camera.Apply();
                aplikacja.ActiveView.Update();
            }
            catch (Exception ex)
            {
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
                if (aplikacja.ActiveDocument != prevActive && prevActive != null)
                {
                    try { prevActive.Activate(); } catch { }
                }
            }

            if (dokumentCzesci.ComponentDefinition == null)
            {
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
                szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[2]);
            }
            catch (Exception ex)
            {
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
                    mf.LogMessage(TranslationManager.GetTranslation("DebugStyleOverride", debugText));
                    break;
                }
            }
            Inventor.TextBox poleTekstowe = null;
            try
            {
                poleTekstowe = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
            }
            catch (Exception ex)
            {
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
            poleTekstowe.FormattedText = formattedText;
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] Pole tekstowe utworzone dla {sciezkaPliku}: {poleTekstowe != null}");
                    break;
                }
            }

            var funkcjeEkstruzji = dokumentCzesci.ComponentDefinition.Features.ExtrudeFeatures;
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
                profil = szkic.Profiles.AddForSolid();
            }
            catch (Exception ex)
            {
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
                operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                    profil,
                    PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                    PartFeatureOperationEnum.kCutOperation
                );
            }
            catch (Exception ex)
            {
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
            operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";
            foreach (Form f in System.Windows.Forms.Application.OpenForms)
            {
                if (f is MainForm mf)
                {
                    mf.LogMessage($"[GRAWEROWANIE] Operacja ekstruzji utworzona dla {sciezkaPliku}: {operacjaEkstrudowaniaNowa != null}");
                    break;
                }
            }

            if ((int)operacjaEkstrudowaniaNowa.HealthStatus == 11780)
            {
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage(TranslationManager.GetTranslation("LogWrongSketchPlane"));
                        break;
                    }
                }
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
                    szkic = dokumentCzesci.ComponentDefinition.Sketches.Add(dokumentCzesci.ComponentDefinition.WorkPlanes[3]);
                }
                catch (Exception ex)
                {
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
                    poleTekstowe2 = szkic.TextBoxes.AddFitted(pozycjaTekstu, tekstGrawerowania);
                }
                catch (Exception ex)
                {
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
                poleTekstowe2.FormattedText = formattedText2;
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
                    nowyProfil = szkic.Profiles.AddForSolid();
                }
                catch (Exception ex)
                {
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
                    operacjaEkstrudowaniaNowa = funkcjeEkstruzji.AddByThroughAllExtent(
                        nowyProfil,
                        PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                        PartFeatureOperationEnum.kCutOperation
                    );
                }
                catch (Exception ex)
                {
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
                operacjaEkstrudowaniaNowa.Name = "Grawerowanie64";
                foreach (Form f in System.Windows.Forms.Application.OpenForms)
                {
                    if (f is MainForm mf)
                    {
                        mf.LogMessage($"[GRAWEROWANIE] Operacja ekstruzji utworzona dla {sciezkaPliku} (druga próba): {operacjaEkstrudowaniaNowa != null}");
                        break;
                    }
                }
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
