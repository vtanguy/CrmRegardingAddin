using System;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Text;

namespace CrmRegardingAddin
{
    public static class CredentialStore
    {
        private const int CRED_TYPE_GENERIC = 1;
        private const int CRED_PERSIST_LOCAL_MACHINE = 2;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CREDENTIAL
        {
            public int Flags;
            public int Type;
            public string TargetName;
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public int CredentialBlobSize;
            public IntPtr CredentialBlob;
            public int Persist;
            public int AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredWrite(ref CREDENTIAL userCredential, uint flags);

        [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CredFree(IntPtr cred);

        [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredDelete(string target, int type, int flags);

        public static string GetDefaultTarget()
        {
            var baseCs = (ConfigurationManager.AppSettings["CrmConnectionBase"] ?? "").Trim();
            var url = ExtractValue(baseCs, "Url") ?? "Default";
            return "CRM_IFD:" + url;
        }

        private static string ExtractValue(string cs, string key)
        {
            if (string.IsNullOrEmpty(cs) || string.IsNullOrEmpty(key)) return null;
            var tok = key + "=";
            var i = cs.IndexOf(tok, StringComparison.OrdinalIgnoreCase);
            if (i < 0) return null;
            var start = i + tok.Length;
            var end = cs.IndexOf(';', start);
            if (end < 0) end = cs.Length;
            return cs.Substring(start, end - start).Trim();
        }

        public static bool Save(string target, string user, string password)
        {
            if (string.IsNullOrEmpty(target) || string.IsNullOrEmpty(user) || password == null) return false;

            var pwdBytes = Encoding.Unicode.GetBytes(password);
            IntPtr blobPtr = Marshal.AllocHGlobal(pwdBytes.Length);
            Marshal.Copy(pwdBytes, 0, blobPtr, pwdBytes.Length);

            var cred = new CREDENTIAL
            {
                AttributeCount = 0,
                Attributes = IntPtr.Zero,
                Comment = null,
                TargetAlias = null,
                Type = CRED_TYPE_GENERIC,
                Persist = CRED_PERSIST_LOCAL_MACHINE,
                CredentialBlobSize = pwdBytes.Length,
                TargetName = target,
                CredentialBlob = blobPtr,
                UserName = user
            };

            bool result = CredWrite(ref cred, 0);
            Marshal.FreeHGlobal(blobPtr);
            return result;
        }

        public static bool TryLoad(string target, out string user, out string password)
        {
            user = null; password = null;
            IntPtr ptr;
            if (!CredRead(target, CRED_TYPE_GENERIC, 0, out ptr) || ptr == IntPtr.Zero)
                return false;

            try
            {
                var cred = (CREDENTIAL)Marshal.PtrToStructure(ptr, typeof(CREDENTIAL));
                user = cred.UserName;
                if (cred.CredentialBlob != IntPtr.Zero && cred.CredentialBlobSize > 0)
                {
                    var bytes = new byte[cred.CredentialBlobSize];
                    Marshal.Copy(cred.CredentialBlob, bytes, 0, bytes.Length);
                    password = Encoding.Unicode.GetString(bytes);
                }
                return !string.IsNullOrEmpty(user) && password != null;
            }
            finally
            {
                CredFree(ptr);
            }
        }

        public static bool Delete(string target)
        {
            return CredDelete(target, CRED_TYPE_GENERIC, 0);
        }
    }
}
