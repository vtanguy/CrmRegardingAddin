// CrmMapiInterop.cs (robust write: string-named + ID-based)
// Target: VS2015 + C#6
using System;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public static class CrmMapiInterop
    {
        // === PS_PUBLIC_STRINGS base ===
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";
        private static string DASL_String(string name) { return "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/" + name; }
        private static string DASL_Id(string hex)     { return "http://schemas.microsoft.com/mapi/id/"     + PS_PUBLIC_STRINGS + "/" + hex; }

        // === Canonical string names (as observed in Microsoft add-in) ===
        // We will write both the *lowercase* and *camelCase* variants to guarantee mapping.
        private const string P_LINKSTATE_LOWER   = "crmlinkstate";
        private const string P_LINKSTATE_CAMEL   = "crmLinkState";
        private const string P_REGID_LOWER       = "crmregardingobjectid";
        private const string P_REGID_CAMEL       = "crmRegardingObjectId";
        private const string P_REGOT_LOWER       = "crmregardingobjecttypecode";
        private const string P_REGOT_CAMEL       = "crmRegardingObjectTypeCode";
        private const string P_CRMID_LOWER       = "crmid";
        private const string P_CRMID_CAMEL       = "crmId";
        private const string P_ORGID_LOWER       = "crmorgid";
        private const string P_ORGID_CAMEL       = "crmOrgId";
        private const string P_SSSPROMOTE        = "crmSssPromoteTracker"; // same in lower/camel in the wild
        private const string P_TRACKEDBY         = "crmTrackedBySender";
        private const string P_OBJTYPE_CAMEL     = "crmObjectTypeCode";    // appt only
        private const string P_ENTRYID_CAMEL     = "crmEntryID";           // appt only

        // === Known ID tags (these may or may not be mapped in a given store) ===
        private static readonly string DASL_crmLinkState   = DASL_Id("0x80C8"); // PT_DOUBLE
        private static readonly string DASL_crmRegardingId = DASL_Id("0x80C9"); // PT_UNICODE
        private static readonly string DASL_crmRegardingOT = DASL_Id("0x80CA"); // PT_UNICODE
        private static readonly string DASL_crmid          = DASL_Id("0x80C4"); // PT_UNICODE
        private static readonly string DASL_crmorgid       = DASL_Id("0x80C5"); // PT_UNICODE
        private static readonly string DASL_crmAsyncSend   = DASL_Id("0x80D7"); // PT_DOUBLE
        private static readonly string DASL_crmSssPromote  = DASL_Id("0x80DE"); // PT_LONG
        private static readonly string DASL_crmTrackedBy   = DASL_Id("0x80DF"); // PT_BOOLEAN
        private static readonly string DASL_crmObjectType  = DASL_Id("0x80D1"); // PT_DOUBLE (appt)
        private static readonly string DASL_crmEntryID     = DASL_Id("0x80C3"); // PT_UNICODE (appt)

        private static readonly string DASL_crmRegardingVT = DASL_Id("0x80CB"); // PT_UNICODE

        // === Safe helpers ===
        private static void TrySet(Outlook.PropertyAccessor pa, string dasl, object value)
        {
            if (pa == null || string.IsNullOrEmpty(dasl)) return;
            try { pa.SetProperty(dasl, value); } catch { }
        }
        private static void TryDelete(Outlook.PropertyAccessor pa, string dasl)
        {
            if (pa == null || string.IsNullOrEmpty(dasl)) return;
            try { pa.DeleteProperty(dasl); } catch { }
        }
        private static void SetBoth(Outlook.PropertyAccessor pa, string idDasl, string nameLower, string nameCamel, object value)
        {
            // Write string-named first (guarantees mapping), then ID (if mapped to expected tag it will succeed)
            TrySet(pa, DASL_String(nameLower), value);
            if (!string.Equals(nameLower, nameCamel, StringComparison.Ordinal))
                TrySet(pa, DASL_String(nameCamel), value);
            TrySet(pa, idDasl, value);
        }
        private static void DeleteBoth(Outlook.PropertyAccessor pa, string idDasl, string nameLower, string nameCamel)
        {
            TryDelete(pa, idDasl);
            TryDelete(pa, DASL_String(nameLower));
            if (!string.Equals(nameLower, nameCamel, StringComparison.Ordinal))
                TryDelete(pa, DASL_String(nameCamel));
        }

        // === Public: remove legacy string-named props that Microsoft used historically ===
        public static void RemoveMsCompatFromMail(Outlook.MailItem mail)
        {
            if (mail == null) return;
            var pa = mail.PropertyAccessor;
            DeleteBoth(pa, DASL_crmLinkState,   P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL);
            DeleteBoth(pa, DASL_crmSssPromote,  P_SSSPROMOTE,      P_SSSPROMOTE);
            DeleteBoth(pa, DASL_crmRegardingId, P_REGID_LOWER,     P_REGID_CAMEL);
            DeleteBoth(pa, DASL_crmRegardingOT, P_REGOT_LOWER,     P_REGOT_CAMEL);
            DeleteBoth(pa, DASL_crmorgid,       P_ORGID_LOWER,     P_ORGID_CAMEL);
            DeleteBoth(pa, DASL_crmid,          P_CRMID_LOWER,     P_CRMID_CAMEL);
            DeleteBoth(pa, DASL_crmTrackedBy,   P_TRACKEDBY,       P_TRACKEDBY);
        }

        public static void RemoveMsCompatFromAppointment(Outlook.AppointmentItem appt)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;
            DeleteBoth(pa, DASL_crmLinkState,   P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL);
            DeleteBoth(pa, DASL_crmSssPromote,  P_SSSPROMOTE,      P_SSSPROMOTE);
            DeleteBoth(pa, DASL_crmRegardingId, P_REGID_LOWER,     P_REGID_CAMEL);
            DeleteBoth(pa, DASL_crmRegardingOT, P_REGOT_LOWER,     P_REGOT_CAMEL);
            DeleteBoth(pa, DASL_crmorgid,       P_ORGID_LOWER,     P_ORGID_CAMEL);
            DeleteBoth(pa, DASL_crmid,          P_CRMID_LOWER,     P_CRMID_CAMEL);
            DeleteBoth(pa, DASL_crmTrackedBy,   P_TRACKEDBY,       P_TRACKEDBY);
            DeleteBoth(pa, DASL_crmObjectType,  P_OBJTYPE_CAMEL,   P_OBJTYPE_CAMEL);
            DeleteBoth(pa, DASL_crmEntryID,     P_ENTRYID_CAMEL,   P_ENTRYID_CAMEL);
        }

        // === Public: set stamps "à la Microsoft" — write string + id
        public static void SetCrmLinkPropsForMail(Outlook.MailItem mi, Guid regardingId, int regardingTypeCode, Guid? crmId, Guid? orgId)
        {
            if (mi == null) return;
            var pa = mi.PropertyAccessor;

            SetBoth(pa, DASL_crmRegardingId, P_REGID_LOWER, P_REGID_CAMEL, "{" + regardingId.ToString().ToUpper() + "}");
            SetBoth(pa, DASL_crmRegardingOT, P_REGOT_LOWER, P_REGOT_CAMEL, regardingTypeCode.ToString());
            SetBoth(pa, DASL_crmSssPromote,  P_SSSPROMOTE,  P_SSSPROMOTE,  1);
            SetBoth(pa, DASL_crmTrackedBy,   P_TRACKEDBY,   P_TRACKEDBY,   false);

            if (crmId.HasValue)
                SetBoth(pa, DASL_crmid,     P_CRMID_LOWER,  P_CRMID_CAMEL,  "{" + crmId.Value.ToString().ToUpper() + "}");
            if (orgId.HasValue)
                SetBoth(pa, DASL_crmorgid,  P_ORGID_LOWER,  P_ORGID_CAMEL,  "{" + orgId.Value.ToString().ToUpper() + "}");

            // LinkState at the end (1.0 soft / 2.0 hard)
            SetBoth(pa, DASL_crmLinkState, P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL, crmId.HasValue ? 2.0 : 1.0);

            try { mi.Save(); } catch { }
        }

        public static void SetCrmLinkPropsForAppointment(Outlook.AppointmentItem appt, Guid regardingId, int regardingTypeCode, Guid? crmId, Guid? orgId, string entryId = null)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;

            SetBoth(pa, DASL_crmRegardingId, P_REGID_LOWER, P_REGID_CAMEL, "{" + regardingId.ToString().ToUpper() + "}");
            SetBoth(pa, DASL_crmRegardingOT, P_REGOT_LOWER, P_REGOT_CAMEL, regardingTypeCode.ToString());
            SetBoth(pa, DASL_crmSssPromote,  P_SSSPROMOTE,  P_SSSPROMOTE,  1);
            SetBoth(pa, DASL_crmTrackedBy,   P_TRACKEDBY,   P_TRACKEDBY,   false);
            SetBoth(pa, DASL_crmObjectType,  P_OBJTYPE_CAMEL, P_OBJTYPE_CAMEL, 4201.0);
            if (!string.IsNullOrEmpty(entryId))
                SetBoth(pa, DASL_crmEntryID, P_ENTRYID_CAMEL, P_ENTRYID_CAMEL, entryId);

            if (crmId.HasValue)
                SetBoth(pa, DASL_crmid,     P_CRMID_LOWER,  P_CRMID_CAMEL,  "{" + crmId.Value.ToString().ToUpper() + "}");
            if (orgId.HasValue)
                SetBoth(pa, DASL_crmorgid,  P_ORGID_LOWER,  P_ORGID_CAMEL,  "{" + orgId.Value.ToString().ToUpper() + "}");

            SetBoth(pa, DASL_crmLinkState, P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL, crmId.HasValue ? 2.0 : 1.0);

            try { appt.Save(); } catch { }
        }

        // === Public: full removal (ID + string) ===
        public static void RemoveCrmLinkProps(Outlook.MailItem mi)
        {
            if (mi == null) return;
            var pa = mi.PropertyAccessor;
            DeleteBoth(pa, DASL_crmLinkState,   P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL);
            DeleteBoth(pa, DASL_crmRegardingId, P_REGID_LOWER,     P_REGID_CAMEL);
            DeleteBoth(pa, DASL_crmRegardingOT, P_REGOT_LOWER,     P_REGOT_CAMEL);
            DeleteBoth(pa, DASL_crmid,          P_CRMID_LOWER,     P_CRMID_CAMEL);
            DeleteBoth(pa, DASL_crmorgid,       P_ORGID_LOWER,     P_ORGID_CAMEL);
            DeleteBoth(pa, DASL_crmAsyncSend,   P_SSSPROMOTE,      P_SSSPROMOTE); // cleanup
            DeleteBoth(pa, DASL_crmSssPromote,  P_SSSPROMOTE,      P_SSSPROMOTE);
            DeleteBoth(pa, DASL_crmTrackedBy,   P_TRACKEDBY,       P_TRACKEDBY);
            try { mi.Save(); } catch { }
        }

        public static void RemoveCrmLinkProps(Outlook.AppointmentItem appt)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;
            DeleteBoth(pa, DASL_crmLinkState,   P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL);
            DeleteBoth(pa, DASL_crmRegardingId, P_REGID_LOWER,     P_REGID_CAMEL);
            DeleteBoth(pa, DASL_crmRegardingOT, P_REGOT_LOWER,     P_REGOT_CAMEL);
            DeleteBoth(pa, DASL_crmid,          P_CRMID_LOWER,     P_CRMID_CAMEL);
            DeleteBoth(pa, DASL_crmorgid,       P_ORGID_LOWER,     P_ORGID_CAMEL);
            DeleteBoth(pa, DASL_crmEntryID,     P_ENTRYID_CAMEL,   P_ENTRYID_CAMEL);
            DeleteBoth(pa, DASL_crmObjectType,  P_OBJTYPE_CAMEL,   P_OBJTYPE_CAMEL);
            DeleteBoth(pa, DASL_crmAsyncSend,   P_SSSPROMOTE,      P_SSSPROMOTE);
            DeleteBoth(pa, DASL_crmSssPromote,  P_SSSPROMOTE,      P_SSSPROMOTE);
            DeleteBoth(pa, DASL_crmTrackedBy,   P_TRACKEDBY,       P_TRACKEDBY);
            try { appt.Save(); } catch { }
        }

        // === Diagnostics ===
        private static string ReadProp(Outlook.PropertyAccessor pa, string dasl)
        {
            if (pa == null || string.IsNullOrEmpty(dasl)) return "(n/a)";
            try { var v = pa.GetProperty(dasl); return v == null ? "(null)" : v.ToString(); }
            catch { return "(absent)"; }
        }
        private static string ReadPropAny(Outlook.PropertyAccessor pa, string idDasl, params string[] names)
        {
            var v = ReadProp(pa, idDasl);
            if (v != "(absent)" && v != "(n/a)") return v;
            for (int i = 0; i < names.Length; i++)
            {
                var vv = ReadProp(pa, DASL_String(names[i]));
                if (vv != "(absent)" && vv != "(n/a)") return vv;
            }
            return v;
        }

        public static string DumpCrmProps(Outlook.MailItem mi)
        {
            if (mi == null) return "Aucun MailItem.";
            var pa = mi.PropertyAccessor;
            var sb = new StringBuilder();
            sb.AppendLine("=== CRM UDF (Mail) — PS_PUBLIC_STRINGS ===");
            sb.AppendLine("crmlinkstate (0x80C8, PT_DOUBLE)       = " + ReadPropAny(pa, DASL_crmLinkState, P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL));
            sb.AppendLine("crmregardingobjectid (0x80C9, PT_UNI)  = " + ReadPropAny(pa, DASL_crmRegardingId, P_REGID_LOWER, P_REGID_CAMEL));
            sb.AppendLine("crmregardingobjecttypecode (0x80CA)    = " + ReadPropAny(pa, DASL_crmRegardingOT, P_REGOT_LOWER, P_REGOT_CAMEL));
            sb.AppendLine("crmid (0x80C4, PT_UNI)                 = " + ReadPropAny(pa, DASL_crmid, P_CRMID_LOWER, P_CRMID_CAMEL));
            sb.AppendLine("crmorgid (0x80C5, PT_UNI)              = " + ReadPropAny(pa, DASL_crmorgid, P_ORGID_LOWER, P_ORGID_CAMEL));
            sb.AppendLine("crmSssPromoteTracker (0x80DE, PT_LONG) = " + ReadPropAny(pa, DASL_crmSssPromote, P_SSSPROMOTE));
            sb.AppendLine("crmTrackedBySender (0x80DF, PT_BOOL)   = " + ReadPropAny(pa, DASL_crmTrackedBy, P_TRACKEDBY));

            sb.AppendLine("regarding (string)                     = " + ReadPropAny(pa, DASL_crmRegardingVT, "Regarding", "regarding"));
            return sb.ToString();
        }

        public static string DumpCrmProps(Outlook.AppointmentItem appt)
        {
            if (appt == null) return "Aucun AppointmentItem.";
            var pa = appt.PropertyAccessor;
            var sb = new StringBuilder();
            sb.AppendLine("=== CRM UDF (Appointment) — PS_PUBLIC_STRINGS ===");
            sb.AppendLine("crmlinkstate (0x80C8, PT_DOUBLE)       = " + ReadPropAny(pa, DASL_crmLinkState, P_LINKSTATE_LOWER, P_LINKSTATE_CAMEL));
            sb.AppendLine("crmregardingobjectid (0x80C9, PT_UNI)  = " + ReadPropAny(pa, DASL_crmRegardingId, P_REGID_LOWER, P_REGID_CAMEL));
            sb.AppendLine("crmregardingobjecttypecode (0x80CA)    = " + ReadPropAny(pa, DASL_crmRegardingOT, P_REGOT_LOWER, P_REGOT_CAMEL));
            sb.AppendLine("crmid (0x80C4, PT_UNI)                 = " + ReadPropAny(pa, DASL_crmid, P_CRMID_LOWER, P_CRMID_CAMEL));
            sb.AppendLine("crmorgid (0x80C5, PT_UNI)              = " + ReadPropAny(pa, DASL_crmorgid, P_ORGID_LOWER, P_ORGID_CAMEL));
            sb.AppendLine("crmObjectTypeCode (0x80D1, PT_DOUBLE)  = " + ReadPropAny(pa, DASL_crmObjectType, P_OBJTYPE_CAMEL));
            sb.AppendLine("crmEntryID (0x80C3, PT_UNI)            = " + ReadPropAny(pa, DASL_crmEntryID, P_ENTRYID_CAMEL));
            sb.AppendLine("crmSssPromoteTracker (0x80DE, PT_LONG) = " + ReadPropAny(pa, DASL_crmSssPromote, P_SSSPROMOTE));
            sb.AppendLine("crmTrackedBySender (0x80DF, PT_BOOL)   = " + ReadPropAny(pa, DASL_crmTrackedBy, P_TRACKEDBY));

            sb.AppendLine("regarding (string)                     = " + ReadPropAny(pa, DASL_crmRegardingVT, "Regarding", "regarding"));
            return sb.ToString();
        }
    }
}
