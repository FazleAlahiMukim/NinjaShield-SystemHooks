using System;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Net.Http;
using Newtonsoft.Json;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMonitor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application outlookApp = new Application();
            outlookApp.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(OnItemSend);

            while (true)
            {
                System.Threading.Thread.Sleep(1000);
            }
        }

        private static void OnItemSend(object Item, ref bool Cancel)
        {
            MailItem mailItem = Item as MailItem;

            if (mailItem != null)
            {
                bool allowSend = CheckAttachmentsWithBackend(mailItem).Result;

                if (!allowSend)
                {
                    Cancel = true;
                }
            }
        }

        private static async Task<bool> CheckAttachmentsWithBackend(MailItem mailItem)
        {
            foreach (Attachment attachment in mailItem.Attachments)
            {
                string tempDirectory = @"C:\Temp";
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }
                string tempPath = Path.Combine(tempDirectory, attachment.FileName);
                attachment.SaveAsFile(tempPath);

                bool allowSend = await CheckWithBackend(tempPath);

                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }

                if (!allowSend)
                {
                    return false;
                }
                
            }

            return true;
        }

        private static async Task<bool> CheckWithBackend(string filePath)
        {
            using (HttpClient client = new HttpClient())
            {
                var request = new { filePath = filePath };
                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("http://localhost:8080/api/checkEmail", httpContent);
                string result = await response.Content.ReadAsStringAsync();

                return result.Equals("OK", StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
