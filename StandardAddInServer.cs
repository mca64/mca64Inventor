using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows.Forms;
using stdole;
using System.Drawing;
using System.IO;

namespace mca64Inventor
{
    /// <summary>
    /// Klasa pomocnicza do konwersji System.Drawing.Image na stdole.IPictureDisp.
    /// Ta klasa dziedziczy po AxHost, aby uzyskać dostęp do chronionej statycznej 
    /// metody GetIPictureDispFromPicture.
    /// </summary>
    [System.ComponentModel.ToolboxItem(false)]
    internal class PictureDispConverter : System.Windows.Forms.AxHost
    {
        // Prywatny konstruktor zapobiega tworzeniu instancji klasy.
        private PictureDispConverter() : base(null)
        {
        }

        /// <summary>
        /// Konwertuje obiekt Image na IPictureDisp.
        /// </summary>
        /// <param name="image">Obraz do konwersji.</param>
        /// <returns>Skonwertowany obiekt IPictureDisp.</returns>
        public static stdole.IPictureDisp GetIPictureDisp(System.Drawing.Image image)
        {
            if (image == null)
            {
                return null;
            }
            // Wywołuje chronioną statyczną metodę z klasy bazowej.
            return GetIPictureDispFromPicture(image) as stdole.IPictureDisp;
        }
    }

    /// <summary>
    /// Jest to główna klasa serwera dodatku, która implementuje interfejs ApplicationAddInServer,
    /// który wszystkie dodatki Inventor muszą implementować. Komunikacja między Inventorem a
    /// dodatkiem odbywa się za pośrednictwem metod tego interfejsu.
    /// </summary>
    [GuidAttribute("a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17")] // <-- To jest NOWY, UNIKALNY GUID dla Twojej wtyczki
    public class AddInServer : ApplicationAddInServer
    {
        // Obiekt aplikacji Inventor.
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_myButton; // Definicja przycisku.

        /// <summary>
        /// Konstruktor klasy AddInServer.
        /// </summary>
        public AddInServer()
        {
        }

        #region ApplicationAddInServer Members

        /// <summary>
        /// Ta metoda jest wywoływana przez Inventora, gdy ładuje dodatek.
        /// Obiekt AddInSiteObject zapewnia dostęp do obiektu aplikacji Inventor.
        /// Flaga FirstTime wskazuje, czy dodatek jest ładowany po raz pierwszy.
        /// </summary>
        /// <param name="addInSiteObject">Obiekt witryny dodatku.</param>
        /// <param name="firstTime">Wskazuje, czy dodatek jest ładowany po raz pierwszy.</param>
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // Inicjalizacja członków dodatku.
            m_inventorApplication = addInSiteObject.Application;

            // Pobiera kolekcję definicji kontrolek.
            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

            // Pobiera wersję zestawu.
            Version assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string versionString = $"v{assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.Build}.{assemblyVersion.Revision}";

            // Tworzy definicję przycisku.
            m_myButton = controlDefs.AddButtonDefinition(
                TranslationManager.GetTranslation("EngraveButton"), // Nazwa wyświetlana z wersją
                "mca64Inventor:GrawerButton", // Nazwa wewnętrzna
                CommandTypesEnum.kShapeEditCmdType,
                "{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}", // ClientId - musi być ten sam GUID co klasy
                TranslationManager.GetTranslation("TooltipEngraveButton"), // Podpowiedź
                TranslationManager.GetTranslation("DescriptionEngraveButton"), // Opis
                PictureDispConverter.GetIPictureDisp(System.Drawing.Image.FromStream(new System.IO.MemoryStream(Resource1.icon16))), // Ikona
                PictureDispConverter.GetIPictureDisp(System.Drawing.Image.FromStream(new System.IO.MemoryStream(Resource1.icon32)))); // Duża ikona

            // Dołącza obsługę zdarzeń.
            m_myButton.OnExecute += OnButtonClick;

            // Definiuje środowiska, w których ma pojawić się zakładka
            string[] environments = { "Part", "Assembly", "Drawing", "ZeroDoc" };

            foreach (string envName in environments)
            {
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons[envName];
                RibbonTab astrusTab;
                try
                {
                    astrusTab = ribbon.RibbonTabs["id_TabAstrus"];
                }
                catch (Exception)
                {
                    // Jeśli zakładka nie istnieje, utwórz ją.
                    astrusTab = ribbon.RibbonTabs.Add("Astrus", "id_TabAstrus", "{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}");
                }
                
                RibbonPanel panel;
                try
                {
                    panel = astrusTab.RibbonPanels["mca64Inventor:GrawerPanel"];
                }
                catch (Exception)
                {
                    // Jeśli panel nie istnieje, utwórz go.
                    panel = astrusTab.RibbonPanels.Add(TranslationManager.GetTranslation("EngravingColumn"), "mca64Inventor:GrawerPanel", "{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}");
                }

                // Dodaje kontrolkę polecenia dla przycisku w panelu.
                panel.CommandControls.AddButton(m_myButton, true);
            }
        }

        /// <summary>
        /// Obsługuje kliknięcie przycisku.
        /// </summary>
        /// <param name="context">Kontekst wywołania.</param>
        private void OnButtonClick(NameValueMap context)
        {
            try
            {
                // Pobiera obiekt aplikacji Inventor.
                Inventor.Application inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;

                if (inventorApp == null)
                {
                    System.Windows.Forms.MessageBox.Show(TranslationManager.GetTranslation("ErrorInventorNotRunningMsgBox"), TranslationManager.GetTranslation("Error"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                if (inventorApp.ActiveDocument == null)
                {
                    System.Windows.Forms.MessageBox.Show(TranslationManager.GetTranslation("ErrorOpenAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                if (!(inventorApp.ActiveDocument is AssemblyDocument))
                {
                    System.Windows.Forms.MessageBox.Show(TranslationManager.GetTranslation("ErrorNotAssemblyMsgBox"), TranslationManager.GetTranslation("Error"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                AssemblyDocument assemblyDoc = (AssemblyDocument)inventorApp.ActiveDocument;

                // Pokazuje MainForm dla aktywnego złożenia.
                MainForm.ShowForAssembly(assemblyDoc.FullFileName);
            }
            catch (Exception ex)
            {
                string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                string logFilePath = System.IO.Path.Combine(desktopPath, "mca64Inventor_ErrorLog.txt");
                string errorMessage = $"Timestamp: {System.DateTime.Now}{System.Environment.NewLine}Error: {ex.Message}{System.Environment.NewLine}{System.Environment.NewLine}Stack Trace:{System.Environment.NewLine}{ex.StackTrace}{System.Environment.NewLine}--------------------------------------------------{System.Environment.NewLine}";
                System.IO.File.AppendAllText(logFilePath, errorMessage);
                System.Windows.Forms.MessageBox.Show(TranslationManager.GetTranslation("ErrorUnexpected", logFilePath), TranslationManager.GetTranslation("Error"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Ta metoda jest wywoływana przez Inventora, gdy dodatek jest wyładowywany.
        /// Dodatek zostanie wyładowany ręcznie przez użytkownika lub
        /// po zakończeniu sesji Inventora.
        /// </summary>
        public void Deactivate()
        {
            // Zwolnienie obiektów.
            m_myButton.OnExecute -= OnButtonClick;
            m_myButton.Delete();
            m_myButton = null;
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Przestarzała metoda, nieużywana.
        /// </summary>
        /// <param name="commandID">ID polecenia.</param>
        public void ExecuteCommand(int commandID)
        {
            // przestarzała metoda, nieużywana
        }

        /// <summary>
        /// Właściwość Automation.
        /// </summary>
        public object Automation
        {
            get { return null; }
        }

        #endregion
    }
}
