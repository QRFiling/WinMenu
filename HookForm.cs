using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WinMenu
{
    class HookForm : Form
    {
        readonly int msgNotify;
        public delegate void EventHandler(IntPtr intPtr);
        public event EventHandler WindowEvent;

        protected virtual void OnWindowEvent(IntPtr intPtr)
        {
            var handler = WindowEvent;
            if (handler != null) handler(intPtr);
        }

        public HookForm()
        {
            msgNotify = Interop.RegisterWindowMessage("SHELLHOOK");
            Interop.RegisterShellHookWindow(Handle);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == msgNotify)
            {
                // Receive shell messages
                switch ((Interop.ShellEvents)m.WParam.ToInt32())
                {
                    case Interop.ShellEvents.HSHELL_WINDOWCREATED:
                        OnWindowEvent(m.LParam);
                        break;
                }
            }
            base.WndProc(ref m);
        }

        protected override void Dispose(bool disposing)
        {
            try { Interop.DeregisterShellHookWindow(this.Handle); }
            catch { }
            base.Dispose(disposing);
        }

        public static class Interop
        {
            public enum ShellEvents : int
            {
                HSHELL_WINDOWCREATED = 1,
                HSHELL_WINDOWDESTROYED = 2,
                HSHELL_ACTIVATESHELLWINDOW = 3,
                HSHELL_WINDOWACTIVATED = 4,
                HSHELL_GETMINRECT = 5,
                HSHELL_REDRAW = 6,
                HSHELL_TASKMAN = 7,
                HSHELL_LANGUAGE = 8,
                HSHELL_ACCESSIBILITYSTATE = 11,
                HSHELL_APPCOMMAND = 12
            }

            [DllImport("user32.dll", EntryPoint = "RegisterWindowMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
            public static extern int RegisterWindowMessage(string lpString);
            [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
            public static extern int DeregisterShellHookWindow(IntPtr hWnd);
            [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
            public static extern int RegisterShellHookWindow(IntPtr hWnd);
        }
    }
}
