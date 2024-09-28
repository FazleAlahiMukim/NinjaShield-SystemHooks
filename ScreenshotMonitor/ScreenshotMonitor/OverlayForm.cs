using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenshotMonitor
{
    public class OverlayForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;


        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

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

        public OverlayForm(int left, int top, int width, int height)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.Black;
            this.Opacity = 1;
            this.TopMost = true;
            this.ShowInTaskbar = false;

            this.StartPosition = FormStartPosition.Manual;
            this.Bounds = new Rectangle(left, top, width, height);

            this.Load += new EventHandler(OverlayForm_Load);
        }

        private void OverlayForm_Load(object sender, EventArgs e)
        {
            this.BringToFront();
            this.Activate();

            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

            Timer closeTimer = new Timer();
            closeTimer.Interval = 5000;
            closeTimer.Tick += (s, ev) =>
            {
                this.Close();
            };
            closeTimer.Start();
        }
    }
}
