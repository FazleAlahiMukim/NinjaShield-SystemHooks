using ScreenshotMonitor;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Automation;
using Tesseract;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.Text;
using System.Threading;
using System.Collections.Generic;
using System.Windows.Threading;


namespace ScreenshotMonitorOcrApp
{
    internal class Program
    {
        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetParent(IntPtr hWnd);

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
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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

        static void AnalyzeScreen()
        {
            List<IntPtr> windows = GetAllVisibleWindows();

            foreach (var hwnd in windows)
            {
                var (bitmap, rect) = CaptureWindow(hwnd);
                if (bitmap == null)
                    continue;


                string extractedText = ExtractTextFromImage(bitmap);
                Task.Run(async () =>
                {
                    bool backendResponse = await CallBackend(extractedText);

                    if (!backendResponse)
                    {
                        Task task = Task.Run(() =>
                        {
                            Rectangle overlayRect = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);

                            using (OverlayForm overlay = new OverlayForm(overlayRect.Left, overlayRect.Top, overlayRect.Width, overlayRect.Height))
                            {
                                overlay.ShowDialog();
                            }
                        });
                    }
                });
            }
        }


        private static bool isSimulatingKeypress = false;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const byte VK_LWIN = 0x5B;  // Left Windows key
        private const byte VK_SHIFT = 0x10; // Shift key
        private const byte VK_S = 0x53;     // 'S' key
        private const uint KEYEVENTF_KEYUP = 0x0002;  // Key up event flag
        private const int VK_ESC = 0x1B;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
        public static extern short GetKeyState(int keyCode);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            

            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                if (vkCode == VK_ESC)
                {
                    StartCooldown(3000);
                }

                bool isWindowsKeyPressed = (GetKeyState(0x5B) & 0x8000) != 0; // 0x5B is for LWin (left Windows key)
                bool isShiftKeyPressed = (GetKeyState(0x10) & 0x8000) != 0;    // 0x10 is for Shift key
                bool isSKeyPressed = (vkCode == 0x53);                         // 0x53 is for 'S' key

                if (isWindowsKeyPressed && isShiftKeyPressed && isSKeyPressed)
                {
                    if (isSimulatingKeypress)
                    {
                        return CallNextHookEx(_hookID, nCode, wParam, lParam);
                    }

                    Task.Run(async () =>
                    {
                        AnalyzeScreen();

                        isSimulatingKeypress = true;
                        SimulateWindowsShiftS();
                        isSimulatingKeypress = false;
                    });
                    

                    
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
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

                var response = await client.PostAsync("http://localhost:8080/api/screenshot", httpContent);
                string result = await response.Content.ReadAsStringAsync();
                return result.Equals("OK", StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void SimulateWindowsShiftS()
        {
            // Press Windows key
            keybd_event(VK_LWIN, 0, 0, UIntPtr.Zero);

            // Press Shift key
            keybd_event(VK_SHIFT, 0, 0, UIntPtr.Zero);

            // Press 'S' key
            keybd_event(VK_S, 0, 0, UIntPtr.Zero);

            // Release 'S' key
            keybd_event(VK_S, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);

            // Release Shift key
            keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);

            // Release Windows key
            keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        static void Main(string[] args)
        {
            SetProcessDPIAware();
            Task.Run(() => MonitorSnippingTool());

            _hookID = SetHook(_proc);
            Application.Run();
            UnhookWindowsHookEx(_hookID);
        }

        public static void MonitorSnippingTool()
        {
            bool eventHandled = false;
            while (true)
            {
                Process[] processes = Process.GetProcessesByName("SnippingTool");
                if (processes.Length > 0)
                {
                    if (!eventHandled)
                    {
                        AutomationElement snippingToolWindow = AutomationElement.FromHandle(processes[0].MainWindowHandle);
                        if (snippingToolWindow != null)
                        {
                            AutomationElement newButton = FindNewButton(snippingToolWindow);
                            AutomationElement cancelButton = FindCancelButton(snippingToolWindow);
                            if (newButton != null && cancelButton != null)
                            {
                                //Console.WriteLine("Found 'New' button in Snipping Tool");
                                SetupEventHandler(newButton, cancelButton);
                                eventHandled = true;
                            }
                        }
                    }
                }
                else
                    eventHandled = false;

                Thread.Sleep(3000);
            }
        }


        private static AutomationElement FindNewButton(AutomationElement root)
        {
            return root.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "New"));
        }

        private static AutomationElement FindCancelButton(AutomationElement root)
        {
            return root.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Cancel"));
        }

        private static bool isCooldown = false;

        private static void SetupEventHandler(AutomationElement newButton, AutomationElement cancelButton)
        {
            AutomationPattern[] patterns = newButton.GetSupportedPatterns();
            foreach (AutomationPattern pattern in patterns)
            {
                if (pattern == InvokePattern.Pattern)
                {
                    Automation.AddAutomationEventHandler(
                        InvokePattern.InvokedEvent,
                        newButton,
                        TreeScope.Element,
                        async (sender, e) =>
                        {
                            if (!isCooldown)
                            {
                                isCooldown = true;
                                await Task.Delay(500);
                                try
                                {
                                    await Task.Run(() =>
                                    {
                                        Dispatcher.CurrentDispatcher.Invoke(() =>
                                        {
                                            InvokePattern invokePattern = (InvokePattern)cancelButton.GetCurrentPattern(InvokePattern.Pattern);
                                            invokePattern.Invoke();
                                        });
                                    });
                                }
                                catch (Exception ex)
                                {
                                }
                                await ContinueWorkflow(newButton);
                                StartCooldown(5000);
                            }

                        });
                    break;
                }
            }
        }

        private static async Task ContinueWorkflow(AutomationElement newButton)
        {
            try
            {
                AnalyzeScreen();

                await Task.Run(() =>
                {
                    Dispatcher.CurrentDispatcher.Invoke(() =>
                    {
                        InvokePattern invokePattern = (InvokePattern)newButton.GetCurrentPattern(InvokePattern.Pattern);
                        invokePattern.Invoke();
                    });
                });
            }
            catch (Exception ex)
            {
            }
        }

        private static void StartCooldown(int n)
        {
            Task.Run(async () =>
            {
                await Task.Delay(n);
                isCooldown = false;
            });
        }

    }
}
