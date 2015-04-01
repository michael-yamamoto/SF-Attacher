using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;

namespace OutlookAddIn2
{
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer = null;
        public Outlook.MAPIFolder selectedFolder = null;
        public Outlook.MailItem currentMailItem = null;

        UserControl1 pane;
        public Tools.CustomTaskPane customPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(currentExplorer_Event);
            
            //creates custom pane on right side of outlook on startup
            pane = new UserControl1();
            customPane = this.CustomTaskPanes.Add(pane, "Settings");
            customPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            customPane.Width = 300;
        }

        void currentExplorer_Event()
        {
            selectedFolder = this.Application.ActiveExplorer().CurrentFolder;

            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        currentMailItem = (selObject as Outlook.MailItem);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
