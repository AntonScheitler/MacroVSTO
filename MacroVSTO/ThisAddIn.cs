using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;
using Microsoft.AspNetCore.WebUtilities;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;

namespace MacroVSTO
{
    public partial class ThisAddIn
    {
        HttpClient client = new HttpClient()
        {
            BaseAddress = new Uri("http://localhost:8080/rest/")
        };

        Dictionary<string, string> parameters = new Dictionary<string, string>()
        {
            ["testExecKey"] = "MOCK-4",
            ["includeTestFields"] = "customfield_10216,customfield_10217,summary"
        };



        private AddInUtilities utilities;

        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
                utilities = new AddInUtilities();

            return utilities;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
