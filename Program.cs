using AutoSMS2.Forms;
using System;
using System.Windows.Forms;

namespace AutoSMS2
{
    public static class Program
    {
        public static bool LoadExcel;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            if (System.Diagnostics.Process
                .GetProcessesByName(
                    System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly()?.Location))
                .Length > 1)
                return;

            bool minimized = false;

            foreach (var arg in args)
            {
                if (arg.Contains("StartMinimized"))
                    minimized = true;
                if (arg.Contains("LoadExcel"))
                    LoadExcel = true;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var mainForm = new Home();

            if (minimized)
            {
                mainForm.WindowState = FormWindowState.Minimized;
                mainForm.ShowIcon = true;
                mainForm.ShowInTaskbar = false;
            }

            Application.Run(mainForm);
        }
    }
}
