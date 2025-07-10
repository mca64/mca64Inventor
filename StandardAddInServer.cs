using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows.Forms;

namespace mca64Inventor
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17")] // <-- To jest NOWY, UNIKALNY GUID dla Twojej wtyczki
    public class AddInServer : ApplicationAddInServer
    {
        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_myButton;

        public AddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application;

            // Get the control definitions collection.
            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

            // Get assembly version
            Version assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string versionString = $"v{assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.Build}.{assemblyVersion.Revision}";

            // Create the button definition.
            m_myButton = controlDefs.AddButtonDefinition(
                $"Uruchom Grawerowanie\n{versionString}", // Display Name with version
                "mca64Inventor:GrawerButton", // Internal Name
                CommandTypesEnum.kShapeEditCmdType,
                "{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}", // ClientId - musi być ten sam GUID co klasy
                "Uruchamia formularz grawerowania.", // Tooltip
                "Przycisk do grawerowania"); // Description

            // Attach the event handler.
            m_myButton.OnExecute += OnButtonClick;

            // Define the environments where the tab should appear
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
                    // If the tab does not exist, create it.
                    astrusTab = ribbon.RibbonTabs.Add("Astrus", "id_TabAstrus", "{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}");
                }
                
                RibbonPanel panel;
                try
                {
                    panel = astrusTab.RibbonPanels["mca64Inventor:GrawerPanel"];
                }
                catch (Exception)
                {
                    // If the panel does not exist, create it.
                    panel = astrusTab.RibbonPanels.Add("Grawerowanie", "mca64Inventor:GrawerPanel", "{a5f3c3e7-7c33-46a2-b16a-42a4a4c25a17}");
                }

                // Add a command control for the button in the panel.
                panel.CommandControls.AddButton(m_myButton, true);
            }
        }

        private void OnButtonClick(NameValueMap context)
        {
            try
            {
                // Get the Inventor application object.
                Inventor.Application inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;

                if (inventorApp == null)
                {
                    System.Windows.Forms.MessageBox.Show("Inventor application is not running.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                if (inventorApp.ActiveDocument == null)
                {
                    System.Windows.Forms.MessageBox.Show("Please open a document in Inventor before running this command.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                if (!(inventorApp.ActiveDocument is AssemblyDocument))
                {
                    System.Windows.Forms.MessageBox.Show("Please open an assembly document in Inventor before running this command.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                AssemblyDocument assemblyDoc = (AssemblyDocument)inventorApp.ActiveDocument;

                // Show the MainForm for the active assembly.
                MainForm.ShowForAssembly(assemblyDoc.FullFileName);
            }
            catch (Exception ex)
            {
                string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                string logFilePath = System.IO.Path.Combine(desktopPath, "mca64Inventor_ErrorLog.txt");
                string errorMessage = $@"Timestamp: {System.DateTime.Now}
Error: {ex.Message}

Stack Trace:
{ex.StackTrace}
--------------------------------------------------
";
                System.IO.File.AppendAllText(logFilePath, errorMessage);
                System.Windows.Forms.MessageBox.Show($@"An unexpected error occurred. 
Details have been written to the log file located at:
{logFilePath}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // Release objects.
            m_myButton.OnExecute -= OnButtonClick;
            m_myButton.Delete();
            m_myButton = null;
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // obsolete method, not used
        }

        public object Automation
        {
            get { return null; }
        }

        #endregion
    }
}
