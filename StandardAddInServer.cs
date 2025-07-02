using System; // Importuje podstawowe klasy .NET, np. do obs�ugi typ�w i wyj�tk�w
using System.Runtime.InteropServices; // Importuje klasy do obs�ugi COM i atrybut�w GUID
using System.Windows.Forms; // Importuje klasy do obs�ugi okienek i komunikat�w
using System.Drawing; // Importuje klasy do obs�ugi grafiki i obraz�w
using Inventor; // Importuje klasy API Autodesk Inventor
using mca64Inventor; // Importuje w�asn� przestrze� nazw projektu

namespace mca64Inventor // Przestrze� nazw grupuj�ca klasy dodatku
{
    /// <summary>
    /// G��wna klasa serwera dodatku Inventor. Implementuje interfejs ApplicationAddInServer.
    /// Odpowiada za rejestracj� przycisku, zak�adki i panelu w interfejsie Inventora.
    /// </summary>
    [GuidAttribute("963308E2-D850-466D-A1C5-503A2E171552")]
    public class AddInServer : Inventor.ApplicationAddInServer
    {
        // Referencja do g��wnej aplikacji Inventor
        private Inventor.Application oInventorApp;
        // Definicja przycisku
        private ButtonDefinition m_ButtonDef;
        // Mened�er interfejsu u�ytkownika
        private UserInterfaceManager m_UIManager;
        // Kategoria polece� dla przycisku
        private CommandCategory m_Category;
        // Zak�adka wst��ki (nieu�ywana bezpo�rednio)
        private RibbonTab m_Tab;
        // Panel wst��ki (nieu�ywany bezpo�rednio)
        private RibbonPanel m_Panel;
        // Sta�e nazwy wewn�trzne dla przycisku, zak�adki i panelu
        private const string ButtonInternalName = "mca64launcherButton";
        private const string TabInternalName = "mca64InventorTab";
        private const string PanelInternalName = "mca64Inventor";

        /// <summary>
        /// Konstruktor domy�lny (nie robi nic specjalnego)
        /// </summary>
        public AddInServer() { }

        /// <summary>
        /// Metoda wywo�ywana przy aktywacji dodatku przez Inventora.
        /// Tworzy zak�adk�, panel i przycisk na wst��ce, je�li nie istniej�.
        /// </summary>
        /// <param name="AddInSiteObject">Obiekt przekazany przez Inventora</param>
        /// <param name="FirstTime">Czy to pierwsza aktywacja</param>
        public void Activate(Inventor.ApplicationAddInSite AddInSiteObject, bool FirstTime)
        {
            oInventorApp = AddInSiteObject.Application; // Pobierz referencj� do aplikacji
            m_UIManager = oInventorApp.UserInterfaceManager; // Pobierz mened�era UI
            try
            {
                // Lista nazw wst��ek, na kt�rych pojawi si� zak�adka
                string[] ribbonNames = { "Part", "Assembly", "Drawing", "ZeroDoc" };
                foreach (var ribbonName in ribbonNames)
                {
                    Ribbon ribbon = m_UIManager.Ribbons[ribbonName]; // Pobierz wst��k�
                    RibbonTab tab = null;
                    // Szukaj istniej�cej zak�adki
                    foreach (RibbonTab t in ribbon.RibbonTabs)
                    {
                        if (t.InternalName == TabInternalName)
                        {
                            tab = t;
                            break;
                        }
                    }
                    // Je�li nie istnieje, utw�rz now� zak�adk�
                    if (tab == null)
                    {
                        tab = ribbon.RibbonTabs.Add("Astrus", TabInternalName, "{963308E2-D850-466D-A1C5-503A2E171552}");
                    }

                    RibbonPanel panel = null;
                    // Szukaj istniej�cego panelu
                    foreach (RibbonPanel p in tab.RibbonPanels)
                    {
                        if (p.InternalName == PanelInternalName)
                        {
                            panel = p;
                            break;
                        }
                    }
                    // Je�li nie istnieje, utw�rz nowy panel
                    if (panel == null)
                    {
                        panel = tab.RibbonPanels.Add("mca64Inventor", PanelInternalName, "{963308E2-D850-466D-A1C5-503A2E171552}");
                    }

                    // Tworzenie przycisku tylko raz
                    if (m_ButtonDef == null)
                    {
                        m_Category = oInventorApp.CommandManager.CommandCategories.Add("mca64Inventor", "{963308E2-D850-466D-A1C5-503A2E171552}");
                        // �adowanie ikon z zasob�w Resource1 (ikony musz� by� dodane do projektu)
                        stdole.IPictureDisp smallIcon = PictureDispConverter.ToIPictureDisp(Resource1.icon16 != null ? (System.Drawing.Image)Image.FromStream(new System.IO.MemoryStream(Resource1.icon16)) : null);
                        stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(Resource1.icon32 != null ? (System.Drawing.Image)Image.FromStream(new System.IO.MemoryStream(Resource1.icon32)) : null);
                        // Dodaj definicj� przycisku do CommandManager
                        m_ButtonDef = oInventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                            "Poka� form�", ButtonInternalName,
                            CommandTypesEnum.kShapeEditCmdType,
                            "{963308E2-D850-466D-A1C5-503A2E171552}",
                            "Wy�wietl przyk�adow� form�64", "Pokazuje okno.", smallIcon, largeIcon, ButtonDisplayEnum.kDisplayTextInLearningMode);
                        // Pod��cz obs�ug� klikni�cia przycisku
                        m_ButtonDef.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(OnButtonExecute);
                    }
                    // Dodaj przycisk do panelu, je�li jeszcze go nie ma
                    bool buttonExists = false;
                    foreach (CommandControl ctrl in panel.CommandControls)
                    {
                        if (ctrl.InternalName == ButtonInternalName)
                        {
                            buttonExists = true;
                            break;
                        }
                    }
                    if (!buttonExists)
                    {
                        panel.CommandControls.AddButton(m_ButtonDef, true);
                    }
                }
            }
            catch (Exception e)
            {
                // Wy�wietl komunikat o b��dzie, je�li co� p�jdzie nie tak
                MessageBox.Show(e.ToString());
            }
        }

        /// <summary>
        /// Metoda wywo�ywana przy dezaktywacji dodatku (np. zamkni�cie Inventora).
        /// Czy�ci referencje i od��cza obs�ug� zdarze�.
        /// </summary>
        public void Deactivate()
        {
            m_ButtonDef.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnButtonExecute); // Od��cz obs�ug� zdarzenia
            m_ButtonDef = null;
            m_UIManager = null;
            m_Category = null;
            m_Tab = null;
            m_Panel = null;
            oInventorApp = null;
            GC.Collect(); // Wymu� oczyszczanie pami�ci
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Obs�uga klikni�cia przycisku na wst��ce. Wy�wietla okno MainForm.
        /// </summary>
        /// <param name="context">Kontekst wywo�ania (nieu�ywany)</param>
        public void OnButtonExecute(NameValueMap context)
        {
            var form = new MainForm(); // Utw�rz nowe okno
            form.ShowDialog(); // Poka� okno jako modalne
        }

        /// <summary>
        /// W�a�ciwo�� automatyzacji (nieu�ywana, ale wymagana przez interfejs)
        /// </summary>
        public object Automation
        {
            get { return null; }
        }

        /// <summary>
        /// Metoda do obs�ugi niestandardowych polece� (opcjonalna, nieu�ywana)
        /// </summary>
        /// <param name="CommandID">Identyfikator polecenia</param>
        public void ExecuteCommand(int CommandID) { /* opcjonalnie: obs�uga niestandardowych polece� */ }
    }
}
