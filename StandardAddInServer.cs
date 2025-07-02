using System; // Importuje podstawowe klasy .NET, np. do obs³ugi typów i wyj¹tków
using System.Runtime.InteropServices; // Importuje klasy do obs³ugi COM i atrybutów GUID
using System.Windows.Forms; // Importuje klasy do obs³ugi okienek i komunikatów
using System.Drawing; // Importuje klasy do obs³ugi grafiki i obrazów
using Inventor; // Importuje klasy API Autodesk Inventor
using mca64Inventor; // Importuje w³asn¹ przestrzeñ nazw projektu

namespace mca64Inventor // Przestrzeñ nazw grupuj¹ca klasy dodatku
{
    /// <summary>
    /// G³ówna klasa serwera dodatku Inventor. Implementuje interfejs ApplicationAddInServer.
    /// Odpowiada za rejestracjê przycisku, zak³adki i panelu w interfejsie Inventora.
    /// </summary>
    [GuidAttribute("963308E2-D850-466D-A1C5-503A2E171552")]
    public class AddInServer : Inventor.ApplicationAddInServer
    {
        // Referencja do g³ównej aplikacji Inventor
        private Inventor.Application oInventorApp;
        // Definicja przycisku
        private ButtonDefinition m_ButtonDef;
        // Mened¿er interfejsu u¿ytkownika
        private UserInterfaceManager m_UIManager;
        // Kategoria poleceñ dla przycisku
        private CommandCategory m_Category;
        // Zak³adka wst¹¿ki (nieu¿ywana bezpoœrednio)
        private RibbonTab m_Tab;
        // Panel wst¹¿ki (nieu¿ywany bezpoœrednio)
        private RibbonPanel m_Panel;
        // Sta³e nazwy wewnêtrzne dla przycisku, zak³adki i panelu
        private const string ButtonInternalName = "mca64launcherButton";
        private const string TabInternalName = "mca64InventorTab";
        private const string PanelInternalName = "mca64Inventor";

        /// <summary>
        /// Konstruktor domyœlny (nie robi nic specjalnego)
        /// </summary>
        public AddInServer() { }

        /// <summary>
        /// Metoda wywo³ywana przy aktywacji dodatku przez Inventora.
        /// Tworzy zak³adkê, panel i przycisk na wst¹¿ce, jeœli nie istniej¹.
        /// </summary>
        /// <param name="AddInSiteObject">Obiekt przekazany przez Inventora</param>
        /// <param name="FirstTime">Czy to pierwsza aktywacja</param>
        public void Activate(Inventor.ApplicationAddInSite AddInSiteObject, bool FirstTime)
        {
            oInventorApp = AddInSiteObject.Application; // Pobierz referencjê do aplikacji
            m_UIManager = oInventorApp.UserInterfaceManager; // Pobierz mened¿era UI
            try
            {
                // Lista nazw wst¹¿ek, na których pojawi siê zak³adka
                string[] ribbonNames = { "Part", "Assembly", "Drawing", "ZeroDoc" };
                foreach (var ribbonName in ribbonNames)
                {
                    Ribbon ribbon = m_UIManager.Ribbons[ribbonName]; // Pobierz wst¹¿kê
                    RibbonTab tab = null;
                    // Szukaj istniej¹cej zak³adki
                    foreach (RibbonTab t in ribbon.RibbonTabs)
                    {
                        if (t.InternalName == TabInternalName)
                        {
                            tab = t;
                            break;
                        }
                    }
                    // Jeœli nie istnieje, utwórz now¹ zak³adkê
                    if (tab == null)
                    {
                        tab = ribbon.RibbonTabs.Add("Astrus", TabInternalName, "{963308E2-D850-466D-A1C5-503A2E171552}");
                    }

                    RibbonPanel panel = null;
                    // Szukaj istniej¹cego panelu
                    foreach (RibbonPanel p in tab.RibbonPanels)
                    {
                        if (p.InternalName == PanelInternalName)
                        {
                            panel = p;
                            break;
                        }
                    }
                    // Jeœli nie istnieje, utwórz nowy panel
                    if (panel == null)
                    {
                        panel = tab.RibbonPanels.Add("mca64Inventor", PanelInternalName, "{963308E2-D850-466D-A1C5-503A2E171552}");
                    }

                    // Tworzenie przycisku tylko raz
                    if (m_ButtonDef == null)
                    {
                        m_Category = oInventorApp.CommandManager.CommandCategories.Add("mca64Inventor", "{963308E2-D850-466D-A1C5-503A2E171552}");
                        // £adowanie ikon z zasobów Resource1 (ikony musz¹ byæ dodane do projektu)
                        stdole.IPictureDisp smallIcon = PictureDispConverter.ToIPictureDisp(Resource1.icon16 != null ? (System.Drawing.Image)Image.FromStream(new System.IO.MemoryStream(Resource1.icon16)) : null);
                        stdole.IPictureDisp largeIcon = PictureDispConverter.ToIPictureDisp(Resource1.icon32 != null ? (System.Drawing.Image)Image.FromStream(new System.IO.MemoryStream(Resource1.icon32)) : null);
                        // Dodaj definicjê przycisku do CommandManager
                        m_ButtonDef = oInventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                            "Poka¿ formê", ButtonInternalName,
                            CommandTypesEnum.kShapeEditCmdType,
                            "{963308E2-D850-466D-A1C5-503A2E171552}",
                            "Wyœwietl przyk³adow¹ formê64", "Pokazuje okno.", smallIcon, largeIcon, ButtonDisplayEnum.kDisplayTextInLearningMode);
                        // Pod³¹cz obs³ugê klikniêcia przycisku
                        m_ButtonDef.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(OnButtonExecute);
                    }
                    // Dodaj przycisk do panelu, jeœli jeszcze go nie ma
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
                // Wyœwietl komunikat o b³êdzie, jeœli coœ pójdzie nie tak
                MessageBox.Show(e.ToString());
            }
        }

        /// <summary>
        /// Metoda wywo³ywana przy dezaktywacji dodatku (np. zamkniêcie Inventora).
        /// Czyœci referencje i od³¹cza obs³ugê zdarzeñ.
        /// </summary>
        public void Deactivate()
        {
            m_ButtonDef.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnButtonExecute); // Od³¹cz obs³ugê zdarzenia
            m_ButtonDef = null;
            m_UIManager = null;
            m_Category = null;
            m_Tab = null;
            m_Panel = null;
            oInventorApp = null;
            GC.Collect(); // Wymuœ oczyszczanie pamiêci
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Obs³uga klikniêcia przycisku na wst¹¿ce. Wyœwietla okno MainForm.
        /// </summary>
        /// <param name="context">Kontekst wywo³ania (nieu¿ywany)</param>
        public void OnButtonExecute(NameValueMap context)
        {
            var form = new MainForm(); // Utwórz nowe okno
            form.ShowDialog(); // Poka¿ okno jako modalne
        }

        /// <summary>
        /// W³aœciwoœæ automatyzacji (nieu¿ywana, ale wymagana przez interfejs)
        /// </summary>
        public object Automation
        {
            get { return null; }
        }

        /// <summary>
        /// Metoda do obs³ugi niestandardowych poleceñ (opcjonalna, nieu¿ywana)
        /// </summary>
        /// <param name="CommandID">Identyfikator polecenia</param>
        public void ExecuteCommand(int CommandID) { /* opcjonalnie: obs³uga niestandardowych poleceñ */ }
    }
}
