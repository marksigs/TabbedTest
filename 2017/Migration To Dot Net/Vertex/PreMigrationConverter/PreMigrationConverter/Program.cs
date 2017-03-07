using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PreMigrationConverter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += 
                new UnhandledExceptionEventHandler(TopLevelErrorHandler);
        
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmMain());
        }

        static void TopLevelErrorHandler(object sender, UnhandledExceptionEventArgs args)
        {
            MessageBox.Show("Error Occured : " + ((Exception)args.ExceptionObject).Message);            
        }
    }
}