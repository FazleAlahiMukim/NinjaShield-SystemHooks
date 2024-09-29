using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using Tesseract;
using System.Net.Http;
using Newtonsoft.Json;


namespace ScreenShareMonitor
{
    internal class Program
    {
        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetParent(IntPtr hWnd);

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);

        static List<IntPtr> GetAllVisibleWindows()
        {
            List<IntPtr> windows = new List<IntPtr>();

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (IsWindowVisible(hWnd) && !IsIconic(hWnd) && !IsSystemWindow(hWnd) &&
                    IsTopLevelWindow(hWnd) && HasWindowTitle(hWnd))
                {
                    windows.Add(hWnd);
                }
                return true;
            }, IntPtr.Zero);

            return windows;
        }

        static bool IsTopLevelWindow(IntPtr hWnd)
        {
            if (GetParent(hWnd) == IntPtr.Zero)
            {
                return true;
            }
            return false;
        }

        static bool HasWindowTitle(IntPtr hWnd)
        {
            const int MaxTitleLength = 256;
            StringBuilder windowText = new StringBuilder(MaxTitleLength);
            GetWindowText(hWnd, windowText, MaxTitleLength);
            if (windowText.Length > 0)
            {
                return true;
            }
            return false;
        }
        static bool IsSystemWindow(IntPtr hWnd)
        {
            const int MaxClassNameLength = 256;
            StringBuilder className = new StringBuilder(MaxClassNameLength);
            GetClassName(hWnd, className, MaxClassNameLength);
            string classNameText = className.ToString();

            if (classNameText == "Progman" || classNameText == "WorkerW" || classNameText == "Shell_TrayWnd" ||
                classNameText == "CabinetWClass" || classNameText == "ExplorerWClass")
            {
                return true;
            }

            return false;
        }

        static (Bitmap, RECT) CaptureWindow(IntPtr hwnd)
        {
            RECT windowRect;
            GetWindowRect(hwnd, out windowRect);

            int width = windowRect.Right - windowRect.Left;
            int height = windowRect.Bottom - windowRect.Top;

            if (width <= 0 || height <= 0)
            {
                return (null, windowRect);
            }

            Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (Graphics gfx = Graphics.FromImage(bmp))
            {
                IntPtr hdc = gfx.GetHdc();
                try
                {
                    PrintWindow(hwnd, hdc, 0);
                }
                finally
                {
                    gfx.ReleaseHdc(hdc);
                }
            }

            return (bmp, windowRect);
        }


        static string ExtractTextFromImage(Bitmap bmp)
        {
            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                using (var img = PixConverter.ToPix(bmp))
                {
                    using (var page = engine.Process(img))
                    {
                        return page.GetText();
                    }
                }
            }
        }

        static bool IsScreenSharingActive()
        {
            string[] screenShareWindows = new string[] {
                "Screen sharing meeting controls"
                // We will add relevant window titles for other softwares
            };

            bool sharing = false;
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                if (IsWindowVisible(hWnd))
                {
                    StringBuilder windowText = new StringBuilder(256);
                    GetWindowText(hWnd, windowText, 256);

                    foreach (var title in screenShareWindows)
                    {
                        if (windowText.ToString().Contains(title))
                        {
                            sharing = true;
                            return false;
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            return sharing;
        }

        static List<Rectangle> activeOverlays = new List<Rectangle>();

        static bool OverlayExists(RECT rect)
        {
            Rectangle newRect = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);

            foreach (var overlayRect in activeOverlays)
            {
                if (overlayRect.Contains(newRect))
                {
                    return true;
                }
            }

            return false;
        }

        public static async Task<bool> CallBackend(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            using (HttpClient client = new HttpClient())
            {
                var request = new { text = text };
                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("http://localhost:8080/api/screenshare", httpContent);
                string result = await response.Content.ReadAsStringAsync();
                return result.Equals("OK", StringComparison.OrdinalIgnoreCase);
            }
        }


        static void Main(string[] args)
        {
            bool wasScreenSharing = false;
            SetProcessDPIAware();

            while (true)
            {
                if (IsScreenSharingActive())
                {
                    List<IntPtr> windows = GetAllVisibleWindows();

                    foreach (var hwnd in windows)
                    {
                        var (bitmap, rect) = CaptureWindow(hwnd);
                        if (bitmap == null)
                            continue;

                        string extractedText = ExtractTextFromImage(bitmap);

                        if (!OverlayExists(rect))
                        {
                            Task.Run(async () =>
                            {
                                bool backendResponse = await CallBackend(extractedText);

                                if (!backendResponse)
                                {
                                    Task task = Task.Run(() =>
                                    {
                                        Rectangle overlayRect = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
                                        activeOverlays.Add(overlayRect);

                                        using (OverlayForm overlay = new OverlayForm(overlayRect.Left, overlayRect.Top, overlayRect.Width, overlayRect.Height))
                                        {
                                            overlay.ShowDialog();
                                        }

                                        activeOverlays.Remove(overlayRect);
                                    });
                                }
                            });
                        }
                    }
                }
                else
                {
                    if (wasScreenSharing)
                    {
                        activeOverlays.Clear();
                        wasScreenSharing = false;
                    }
                }

                System.Threading.Thread.Sleep(3000);
            }
        }
    }
}
