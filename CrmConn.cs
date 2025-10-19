using System;
using System.Configuration;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
// Correct namespace for WhoAmI*:
using Microsoft.Crm.Sdk.Messages;

namespace CrmRegardingAddin
{
    public static class CrmConn
    {
        public static IOrganizationService ConnectWithCredentials(string user, string pass, out string diag)
        {
            diag = null;
            try
            {
                try { ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; } catch { }

                var baseCs = (ConfigurationManager.AppSettings["CrmConnectionBase"] ?? "").Trim();
                if (string.IsNullOrEmpty(baseCs))
                {
                    baseCs = (ConfigurationManager.AppSettings["CrmConnectionString"] ?? "").Trim();
                }
                if (string.IsNullOrEmpty(baseCs))
                {
                    diag = "App.config: missing 'CrmConnectionBase' (or 'CrmConnectionString').";
                    return null;
                }

                string cleaned = RemoveKey(baseCs, "Username");
                cleaned = RemoveKey(cleaned, "UserName");
                cleaned = RemoveKey(cleaned, "Password");
                cleaned = RemoveKey(cleaned, "Integrated Security");
                cleaned = EnsureKey(cleaned, "RequireNewInstance", "true");

                if (cleaned.IndexOf("{USERNAME}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    cleaned.IndexOf("{PASSWORD}", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cleaned = cleaned.Replace("{USERNAME}", user).Replace("{PASSWORD}", pass);
                }
                else
                {
                    if (!HasKey(cleaned, "Username")) cleaned += (cleaned.EndsWith(";") ? "" : ";") + "Username=" + user + ";";
                    if (!HasKey(cleaned, "Password")) cleaned += "Password=" + pass + ";";
                }

                if (IsIFD(cleaned) && !HasKey(cleaned, "SkipDiscovery"))
                    cleaned += "SkipDiscovery=true;";

                var csc = new CrmServiceClient(cleaned);
                if (csc == null || !csc.IsReady)
                {
                    diag = "LastCrmError: " + (csc != null ? csc.LastCrmError : "(CrmServiceClient=null)");
                    if (csc != null && csc.LastCrmException != null) diag += Environment.NewLine + "LastException: " + csc.LastCrmException;
                    return null;
                }

                IOrganizationService svc = null;
                if (csc.OrganizationWebProxyClient != null)
                    svc = (IOrganizationService)csc.OrganizationWebProxyClient;
                else if (csc.OrganizationServiceProxy != null)
                    svc = (IOrganizationService)csc.OrganizationServiceProxy;

                if (svc == null)
                {
                    diag = "IsReady==true but no organization proxy available.";
                    return null;
                }

                try
                {
                    var resp = (WhoAmIResponse)svc.Execute(new WhoAmIRequest());
                    if (resp == null || resp.UserId == Guid.Empty)
                    {
                        diag = "Connected but WhoAmI returned empty UserId.";
                        return null;
                    }
                }
                catch (Exception exWho)
                {
                    diag = "WhoAmI failed: " + exWho.Message;
                    return null;
                }

                return svc;
            }
            catch (Exception ex)
            {
                diag = "EX: " + ex.Message + (ex.InnerException != null ? ("\r\nINNER: " + ex.InnerException.Message) : "");
                return null;
            }
        }

        private static bool IsIFD(string cs)
        {
            var kv = GetValue(cs, "AuthType");
            return !string.IsNullOrEmpty(kv) && kv.Trim().Equals("IFD", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasKey(string cs, string key)
        {
            return GetValue(cs, key) != null;
        }

        private static string EnsureKey(string cs, string key, string value)
        {
            return HasKey(cs, key) ? cs : (cs.TrimEnd().TrimEnd(';') + ";" + key + "=" + value + ";");
        }

        private static string RemoveKey(string cs, string key)
        {
            int idx = IndexOfKey(cs, key);
            while (idx >= 0)
            {
                int start = idx;
                int end = cs.IndexOf(';', start);
                if (end < 0) end = cs.Length;
                cs = cs.Remove(start, end - start);
                idx = IndexOfKey(cs, key);
            }
            return cs;
        }

        private static int IndexOfKey(string cs, string key)
        {
            var pattern = key + "=";
            int i = cs.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
            if (i < 0) return -1;
            int start = i;
            while (start > 0 && cs[start - 1] != ';') start--;
            return start;
        }

        private static string GetValue(string cs, string key)
        {
            var pattern = key + "=";
            int i = cs.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
            if (i < 0) return null;
            int start = i + pattern.Length;
            int end = cs.IndexOf(';', start);
            if (end < 0) end = cs.Length;
            return cs.Substring(start, end - start).Trim();
        }
    }
}
