using System;
using System.IO;
using System.Configuration;

namespace CrmRegardingAddin
{
    public static class Logger
    {
        private static readonly bool _enabled;
        private static readonly string _logPath;

        static Logger()
        {
            try
            {
                var enabled = ConfigurationManager.AppSettings["CrmAddin:VerboseLog"];
                _enabled = string.Equals(enabled, "true", StringComparison.OrdinalIgnoreCase);

                var configuredPath = ConfigurationManager.AppSettings["CrmAddin:LogFile"];
                if (string.IsNullOrWhiteSpace(configuredPath))
                {
                    configuredPath = @"%LOCALAPPDATA%\CrmRegardingAddin\crmaddin.log";
                }

                var expanded = Environment.ExpandEnvironmentVariables(configuredPath);
                var dir = Path.GetDirectoryName(expanded);
                if (!string.IsNullOrEmpty(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                _logPath = expanded;
            }
            catch
            {
                _enabled = false;
                _logPath = null;
            }
        }

        public static void Info(string message)
        {
            if (!_enabled || string.IsNullOrEmpty(_logPath)) return;
            try
            {
                File.AppendAllText(_logPath, DateTime.Now.ToString("u") + " " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
