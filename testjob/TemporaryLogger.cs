using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace testjob
{
    public class TemporaryLogger
    {
        private readonly string logFilePath;
        private readonly StreamWriter streamWriter;

        public TemporaryLogger()
        {
            string logDirectory = Path.Combine(Path.GetTempPath(), "ConsoleLogs");
            Directory.CreateDirectory(logDirectory);

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            logFilePath = Path.Combine(logDirectory, $"TempLog_{timestamp}.txt");

            streamWriter = new StreamWriter(logFilePath, true, Encoding.UTF8);
        }

        public void LogMessage(string message)
        {
            string logEntry = $"{DateTime.Now}: {message}";
            Console.WriteLine(logEntry);
            streamWriter.WriteLine(logEntry);
            streamWriter.Flush();
        }

        public void DisplayLogInNotepad()
        {
            streamWriter.Flush();
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "notepad.exe",
                    Arguments = logFilePath,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Normal
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error opening Notepad: " + ex.Message);
            }
        }

        public void Dispose()
        {
            streamWriter?.Dispose();
        }
    }
}
