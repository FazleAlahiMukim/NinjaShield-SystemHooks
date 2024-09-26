using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Net.Http;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

namespace PrintJobMonitor
{
    internal class Program
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetJob(IntPtr hPrinter, int JobID, int Level, IntPtr pJob, int Command);

        [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

        [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        const int JOB_CONTROL_PAUSE = 1;
        const int JOB_CONTROL_RESUME = 2;
        const int JOB_CONTROL_DELETE = 5;

        static void Main(string[] args)
        {
            ManagementEventWatcher watcher = new ManagementEventWatcher();
            watcher.Query = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
            watcher.EventArrived += new EventArrivedEventHandler(PrintJobArrived);
            watcher.Start();
            while (true)
            {
                // Keep monitoring the print jobs
                System.Threading.Thread.Sleep(1000);
            }
        }

        static void PrintJobArrived(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject printJob = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            string documentName = printJob["Document"].ToString();
            string printJobId = printJob["JobId"].ToString();
            string printerName = printJob["Name"].ToString();
            string[] nameParts = printerName.Split(',');
            printerName = nameParts[0].Trim();

            // Pause the print job
            PausePrintJob(printerName, int.Parse(printJobId));

            // Send file path to backend and wait for approval
            bool allowPrint = CheckWithBackend(documentName).GetAwaiter().GetResult();

            if (allowPrint)
            {
                // Resume the print job if backend says OK
                ResumePrintJob(printerName, int.Parse(printJobId));
            }
            else
            {
                // Delete the print job
                DeletePrintJob(printerName, int.Parse(printJobId));
            }
        }

        private static async Task<bool> CheckWithBackend(string documentPath)
        {
            using (HttpClient client = new HttpClient())
            {
                var request = new { filePath = documentPath };
                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("http://localhost:8080/api/printJob", httpContent);
                string result = await response.Content.ReadAsStringAsync();
                return result.Equals("OK", StringComparison.OrdinalIgnoreCase);
            }
        }

        public static void PausePrintJob(string printerName, int jobId)
        {
            IntPtr hPrinter;
            if (OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
            {
                SetJob(hPrinter, jobId, 0, IntPtr.Zero, JOB_CONTROL_PAUSE);
                ClosePrinter(hPrinter);
            }
        }

        public static void ResumePrintJob(string printerName, int jobId)
        {
            IntPtr hPrinter;
            if (OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
            {
                SetJob(hPrinter, jobId, 0, IntPtr.Zero, JOB_CONTROL_RESUME);
                ClosePrinter(hPrinter);
            }
        }

        public static void DeletePrintJob(string printerName, int jobId)
        {
            IntPtr hPrinter;
            if (OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
            {
                SetJob(hPrinter, jobId, 0, IntPtr.Zero, JOB_CONTROL_DELETE);
                ClosePrinter(hPrinter);
            }
        }

    }
}
