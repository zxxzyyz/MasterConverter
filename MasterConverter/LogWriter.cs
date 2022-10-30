using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Text.RegularExpressions;

namespace MasterConverter
{
    static class LogWriter
    {
        private static string fileNameLog = Assembly.GetExecutingAssembly().GetName().Name + "Log";

        private static string fileNameErrorLog = Assembly.GetExecutingAssembly().GetName().Name + "ErrorLog";

        private static Regex regexLog = new Regex(Assembly.GetExecutingAssembly().GetName().Name + "Log_(\\d{8}).txt");

        private static Regex regexErrorLog = new Regex(Assembly.GetExecutingAssembly().GetName().Name + "ErrorLog_(\\d{8}).txt");

        private static string logPath = $"{AppDomain.CurrentDomain.BaseDirectory}{Properties.Settings.Default["LogFolder"]}";

        public static void WriteLog(string file, string newFile, string content)
        {
            var sb = new StringBuilder();
            var fileName = Path.GetFileName(file);
            var newFileName = Path.GetFileName(newFile);
            var logFile = $"{logPath}\\{fileNameLog}_{DateTime.Now.ToString("yyyyMMdd")}.txt";

            CreateFolder(logFile);

            using (StreamWriter writer = new StreamWriter(logFile, true))
            {
                sb.Append(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                sb.Append($" {fileName} -> {newFileName} ({content})");
                writer.WriteLine(sb.ToString());
            }
        }

        public static void WriteErrorLog(string log)
        {
            var logFile = $"{logPath}\\{fileNameErrorLog}_{DateTime.Now.ToString("yyyyMMdd")}.txt";
            CreateFolder(logFile);
            using (StreamWriter writer = new StreamWriter(logFile, true))
            {
                writer.WriteLine(log);
            }
        }

        public static void DeleteLog()
        {
            CreateFolder($"{logPath}\\");

            string[] files = Directory.GetFiles(logPath);

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                string fileName = Path.GetFileName(fi.FullName);

                if (regexLog.IsMatch(fileName) || regexErrorLog.IsMatch(fileName))
                {
                    if (fi.CreationTime < DateTime.Now.AddDays(-int.Parse(Properties.Settings.Default["LogSavePeriod"].ToString())))
                    {
                        fi.Delete();
                    }
                }
            }
        }

        private static void CreateFolder(string filePath)
        {
            FileInfo fi = new FileInfo(filePath);
            if (!fi.Directory.Exists)
            {
                Directory.CreateDirectory(fi.DirectoryName);
            }
        }
    }
}
