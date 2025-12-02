using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Documentupdate
{
    public static class Loggerfile
    {
        private static readonly string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");

        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }

        public static void LogError(string message, Exception ex = null)
        {
            Log("ERROR", $"{message} {(ex != null ? ex.ToString() : "")}");
        }

        private static void Log(string level, string message)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {level}: {message}";
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error logging message: {ex.Message}");
            }
        }
    }
}
