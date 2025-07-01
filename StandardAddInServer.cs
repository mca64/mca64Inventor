using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using mca64Inventor; // Zmieniono nazwê przestrzeni nazw

namespace mca64Inventor
{
    [GuidAttribute("963308E2-D850-466D-A1C5-503A2E171552")]
    public class AddInServer : Inventor.ApplicationAddInServer
    {
        private Inventor.Application oInventorApp;
        private ButtonDefinition m_ButtonDef;
        private UserInterfaceManager m_UIManager;
        private CommandCategory m_Category;
        private RibbonTab m_Tab;
        private RibbonPanel m_Panel;
        private const string ButtonInternalName = "mca64launcherButton";
        private const string TabInternalName = "mca64InventorTab";
        private const string PanelInternalName = "mca64InventorPanel";

        public AddInServer() { }

        public void Activate(Inventor.ApplicationAddInSite AddInSiteObject, bool FirstTime)
        {
            oInventorApp = AddInSiteObject.Application;
            m_UIManager = oInventorApp.UserInterfaceManager;
            try
            {
                string[] ribbonNames = { "Part", "Assembly", "Drawing", "ZeroDoc" };
                foreach (var ribbonName in ribbonNames)
                {
                    Ribbon ribbon = m_UIManager.Ribbons[ribbonName];
                    RibbonTab tab = null;
                    foreach (RibbonTab t in ribbon.RibbonTabs)
                    {
                        if (t.InternalName == TabInternalName)
                        {
                            tab = t;
                            break;
                        }
                    }
                    if (tab == null)
                    {
                        tab = ribbon.RibbonTabs.Add("mca64Inventor", TabInternalName, "{963308E2-D850-466D-A1C5-503A2E171552}");
                    }

                    RibbonPanel panel = null;
                    foreach (RibbonPanel p in tab.RibbonPanels)
                    {
                        if (p.InternalName == PanelInternalName)
                        {
                            panel = p;
                            break;
                        }
                    }
                    if (panel == null)
                    {
                        panel = tab.RibbonPanels.Add("mca64InventorPanel", PanelInternalName, "{963308E2-D850-466D-A1C5-503A2E171552}");
                    }

                    if (m_ButtonDef == null)
                    {
                        m_Category = oInventorApp.CommandManager.CommandCategories.Add("mca64Inventor", "{963308E2-D850-466D-A1C5-503A2E171552}");
                        m_ButtonDef = oInventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                            "Poka¿ formê", ButtonInternalName,
                            CommandTypesEnum.kShapeEditCmdType,
                            "{963308E2-D850-466D-A1C5-503A2E171552}",
                            "Wyœwietl przyk³adow¹ formê64", "Pokazuje okno.", null, null, ButtonDisplayEnum.kDisplayTextInLearningMode);
                        m_ButtonDef.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(OnButtonExecute);
                    }
                    // Dodaj przycisk tylko jeœli nie istnieje
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
                MessageBox.Show(e.ToString());
            }
        }

        public void Deactivate()
        {
            m_ButtonDef.OnExecute -= new ButtonDefinitionSink_OnExecuteEventHandler(OnButtonExecute);
            m_ButtonDef = null;
            m_UIManager = null;
            m_Category = null;
            m_Tab = null;
            m_Panel = null;
            oInventorApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnButtonExecute(NameValueMap context)
        {
            var form = new MainForm();
            form.ShowDialog();
        }

        public object Automation
        {
            get { return null; }
        }

        public void ExecuteCommand(int CommandID) { /* opcjonalnie: obs³uga niestandardowych poleceñ */ }
    }
}
