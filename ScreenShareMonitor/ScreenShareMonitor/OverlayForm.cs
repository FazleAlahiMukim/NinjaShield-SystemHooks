using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ScreenShareMonitor
{
    public class OverlayForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;

        public OverlayForm(int left, int top, int width, int height)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.Red;
            this.Opacity = 0.5;
            this.TopMost = true;
            this.ShowInTaskbar = false;

            this.StartPosition = FormStartPosition.Manual;
            this.Bounds = new Rectangle(left, top, width, height);

            this.Load += new EventHandler(OverlayForm_Load);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            this.Close();
        }

        private void OverlayForm_Load(object sender, EventArgs e)
        {
            this.BringToFront();
            this.Activate();

            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        }
    }
}
