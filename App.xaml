using System;
using System.IO;
using System.Windows;

namespace KarzaConsolidator
{
    public partial class App : Application
    {
        public App()
        {
            // Catch UI thread rendering exceptions
            this.DispatcherUnhandledException += (s, e) =>
            {
                LogCrash("UI_Exception", e.Exception);
                e.Handled = true;
            };

            // Catch background initialization exceptions
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                LogCrash("Background_Exception", e.ExceptionObject as Exception);
            };
        }

        private void LogCrash(string type, Exception ex)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "KARZA_CRASH_LOG.txt");
                string errorMsg = $"[{DateTime.Now}] {type}:\r\n{ex?.Message}\r\n{ex?.StackTrace}\r\n";
                
                if (ex?.InnerException != null)
                {
                    errorMsg += $"Inner Exception: {ex.InnerException.Message}\r\n{ex.InnerException.StackTrace}\r\n";
                }
                
                File.AppendAllText(logPath, errorMsg);
                
                MessageBox.Show($"The engine encountered a critical startup block.\n\nA crash log has been saved to:\n{logPath}\n\nError: {ex?.Message}", "System Initialization Failure", MessageBoxButton.OK, MessageBoxImage.Error);
                Environment.Exit(1);
            }
            catch 
            { 
                Environment.Exit(1); 
            }
        }
    }
}
