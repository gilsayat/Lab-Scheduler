using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace LabtechScheduler
{
    static class Program
    {
        private static Mutex mutex = null;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private const int SW_RESTORE = 9;

        [STAThread]
        static void Main()
        {
            const string appName = "LabtechScheduler_v1.7_SingleInstance_Mutex";
            bool createdNew;

            mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                IntPtr existingWindow = FindWindow(null, "Labtech Scheduler v1.7");
                if (existingWindow != IntPtr.Zero)
                {
                    ShowWindow(existingWindow, SW_RESTORE);
                    SetForegroundWindow(existingWindow);
                }
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());

            GC.KeepAlive(mutex);
        }
    }
}
