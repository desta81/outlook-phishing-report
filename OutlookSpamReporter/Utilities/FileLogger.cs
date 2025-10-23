using System;
using System.Diagnostics;
using System.IO;

namespace OutlookSpamReporter.Utilities
{
    public static class FileLogger
    {
        private static readonly object SyncLock = new object();
        private static readonly string LogDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookSpamReporter", "Logs");
        private static readonly string LogFilePath = Path.Combine(LogDirectoryPath, "add-in.log");

        public static void Info(string message)
        {
            Write("INFO", message);
        }

        public static void Error(string message, Exception ex = null)
        {
            string full = ex == null ? message : message + " | Exception: " + ex;
            Write("ERROR", full);
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (SyncLock)
                {
                    if (!Directory.Exists(LogDirectoryPath))
                    {
                        Directory.CreateDirectory(LogDirectoryPath);
                    }
                    string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " [" + level + "] " + message;
                    File.AppendAllText(LogFilePath, line + Environment.NewLine);
                }
            }
            catch
            {
                try
                {
                    Debug.WriteLine("Logging failed: " + level + ": " + message);
                }
                catch { }
            }
        }
    }
}


