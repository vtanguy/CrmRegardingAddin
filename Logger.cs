using System.Diagnostics;

namespace CrmRegardingAddin
{
    internal static class Logger
    {
        public static void Info(string message)
        {
            Debug.WriteLine("[CRMADDIN] " + message);
        }
    }
}
