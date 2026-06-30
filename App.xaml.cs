using System;
using System.IO;
using System.Windows;

namespace KarzaConsolidator
{
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                LogCrashDump("UI_Thread_Exception", e.Exception);
                e.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                LogCrashDump("AppDomain_Core_Exception", e.ExceptionObject as Exception);
            };
        }

        private void LogCrashDump(string errorContext, Exception? exception)
        {
            try
            {
                string executionFolder = AppDomain.CurrentDomain.BaseDirectory;
                string dumpFilePath = Path.Combine(executionFolder, "KARZA_CRASH_LOG.txt");
                
                string failurePayload = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Context: {errorContext}\r\n" +
                                        $"Message: {exception?.Message}\r\n" +
                                        $"StackTrace:\r\n{exception?.StackTrace}\r\n";

                if (exception?.InnerException != null)
                {
                    failurePayload += $"Inner Matrix Error: {exception.InnerException.Message}\r\n" +
                                     $"Inner StackTrace:\r\n{exception.InnerException.StackTrace}\r\n";
                }

                File.AppendAllText(dumpFilePath, failurePayload);
                
                MessageBox.Show($"Critical framework disruption captured by safety layer.\n\nCrash file compiled at:\n{dumpFilePath}\n\nError Vector: {exception?.Message}", 
                                "Engine Failure", MessageBoxButton.OK, MessageBoxImage.Error);
                
                Environment.Exit(1);
            }
            catch
            {
                Environment.Exit(1);
            }
        }
    }
}
