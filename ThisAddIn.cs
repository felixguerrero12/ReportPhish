using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace PhishingReportAddin
{
    public partial class ThisAddIn
    {
        private PhishingRibbon ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("PhishingReport Add-in is starting up");
            ribbon = new PhishingRibbon();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            // must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            System.Diagnostics.Debug.WriteLine("CreateRibbonExtensibilityObject called");
            ribbon = new PhishingRibbon();
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
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