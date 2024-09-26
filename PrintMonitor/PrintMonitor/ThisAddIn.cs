using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Net.Http;
using Newtonsoft.Json;
using System.Threading.Tasks;

namespace PrintMonitor
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentBeforePrint += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforePrintEventHandler(Application_DocumentBeforePrint);
        }

        private void Application_DocumentBeforePrint(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            string documentPath = Doc.FullName;

            Cancel = true;

            bool allowPrint = CheckWithBackend(documentPath).GetAwaiter().GetResult();

            if (allowPrint)
            {
                Doc.PrintOut();
            }
        }

        private async Task<bool> CheckWithBackend(string documentPath)
        {
            using (HttpClient client = new HttpClient())
            {
                var request = new { filePath = documentPath };
                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("http://localhost:8080/api/printWord", httpContent);
                string result = await response.Content.ReadAsStringAsync();
                return result.Equals("OK", StringComparison.OrdinalIgnoreCase);
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
