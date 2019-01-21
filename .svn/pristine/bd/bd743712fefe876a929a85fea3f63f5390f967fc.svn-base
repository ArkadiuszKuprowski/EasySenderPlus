using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace easyDMSTool
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.Application outlook = Globals.ThisAddIn.Application;
            outlook.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);            
            CreateRibbonExtensibilityObject();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {          
            //ribbonEasyDMS.file.Close();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ribbonEasyDMS();
        }
        void Application_OptionsPagesAdd(Microsoft.Office.Interop.Outlook.PropertyPages Pages)
        {
            Pages.Add(new easyDMSToolOptionDialog(), "");
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
