using System;
using System.IO;
using System.Configuration;

namespace CrmRegardingAddin
{
    public static class Logger
    {
        private static bool _enabled;
        private static string _logPath;

        static Logger()
        {
            _enabled = ConfigurationManager.AppSettings["CrmAddin:VerboseLog"] == "true";
            _logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CrmRegardingAddin", "crmaddin.log");
            Directory.CreateDirectory(Path.GetDirectoryName(_logPath));
        }

        public static void Info(string message)
        {
            if (!_enabled) return;
            try
            {
                File.AppendAllText(_logPath, DateTime.Now.ToString("u") + " " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
