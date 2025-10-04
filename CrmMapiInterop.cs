
// CrmMapiInterop.cs
// VS2015 + C#6 compatible
using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public static class CrmMapiInterop
    {
        private const string DASL_BASE = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

        private const string P_LINKSTATE   = "crmLinkState";
        private const string P_SSS_TRACK   = "crmSssPromoteTracker";
        private const string P_REGARDINGID = "crmRegardingId";
        private const string P_REGARDINGOT = "crmRegardingObjectType";
        private const string P_REGARDING   = "Regarding";
        private const string P_ORGID       = "crmorgid";
        private const string P_PARTYINFO   = "crmpartyinfo";

        private static string N(string name) { return DASL_BASE + name; }

        public static void ApplyMsCompatForMail(
            Outlook.MailItem mail,
            string regardingLogicalNameOrCode,
            Guid regardingId,
            string regardingDisplayName,
            string systemUserEmail,
            Guid? systemUserId,
            string fromSmtp,
            IEnumerable<string> recipients,
            bool isIncoming,
            Guid orgId
        )
        {
            if (mail == null) return;

            var pa = mail.PropertyAccessor;

            TrySet(pa, N(P_LINKSTATE), 2.0);
            TrySet(pa, N(P_SSS_TRACK), 0);

            if (regardingId != Guid.Empty)
                TrySet(pa, N(P_REGARDINGID), ToBracedUpper(regardingId));

            int typeCode;
            if (!string.IsNullOrEmpty(regardingLogicalNameOrCode) && int.TryParse(regardingLogicalNameOrCode, out typeCode))
                TrySet(pa, N(P_REGARDINGOT), typeCode.ToString());

            if (!string.IsNullOrEmpty(regardingDisplayName))
                TrySet(pa, N(P_REGARDING), regardingDisplayName);

            if (orgId != Guid.Empty)
                TrySet(pa, N(P_ORGID), ToBracedUpper(orgId));

            string xml = BuildCrmPartyInfoXmlForMail(systemUserEmail, systemUserId, fromSmtp, recipients, isIncoming);
            if (!string.IsNullOrEmpty(xml))
                TrySet(pa, N(P_PARTYINFO), xml);

            try { mail.Save(); } catch { }
        }

        public static void RemoveMsCompatFromMail(Outlook.MailItem mail)
        {
            if (mail == null) return;
            var pa = mail.PropertyAccessor;
            TryDelete(pa, N(P_LINKSTATE));
            TryDelete(pa, N(P_SSS_TRACK));
            TryDelete(pa, N(P_REGARDINGID));
            TryDelete(pa, N(P_REGARDINGOT));
            TryDelete(pa, N(P_REGARDING));
            TryDelete(pa, N(P_ORGID));
            TryDelete(pa, N(P_PARTYINFO));
            try { mail.Save(); } catch { }
        }

        public static void ApplyMsCompatForAppointment(
            Outlook.AppointmentItem appt,
            Guid regardingId,
            string regardingDisplayName,
            string organizerSmtp,
            Guid? organizerSystemUserId
        )
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;

            TrySet(pa, N(P_LINKSTATE), 2.0);
            TrySet(pa, N(P_SSS_TRACK), 0);

            if (regardingId != Guid.Empty)
                TrySet(pa, N(P_REGARDINGID), ToBracedUpper(regardingId));

            if (!string.IsNullOrEmpty(regardingDisplayName))
                TrySet(pa, N(P_REGARDING), regardingDisplayName);

            string xml = BuildCrmPartyInfoXmlForAppointment(organizerSmtp, organizerSystemUserId);
            if (!string.IsNullOrEmpty(xml))
                TrySet(pa, N(P_PARTYINFO), xml);

            try { appt.Save(); } catch { }
        }

        public static void RemoveMsCompatFromAppointment(Outlook.AppointmentItem appt)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;
            TryDelete(pa, N(P_LINKSTATE));
            TryDelete(pa, N(P_SSS_TRACK));
            TryDelete(pa, N(P_REGARDINGID));
            TryDelete(pa, N(P_REGARDINGOT));
            TryDelete(pa, N(P_REGARDING));
            TryDelete(pa, N(P_PARTYINFO));
            try { appt.Save(); } catch { }
        }

        private static void TrySet(Outlook.PropertyAccessor pa, string dasl, object value)
        {
            try { pa.SetProperty(dasl, value); } catch { }
        }

        private static void TryDelete(Outlook.PropertyAccessor pa, string dasl)
        {
            try { pa.DeleteProperty(dasl); } catch { }
        }

        private static string ToBracedUpper(Guid id)
        {
            return "{" + id.ToString().ToUpper() + "}";
        }

        private static string X(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private static string BuildCrmPartyInfoXmlForMail(
            string systemUserEmail,
            Guid? systemUserId,
            string fromSmtp,
            IEnumerable<string> recipients,
            bool isIncoming)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("<PartyMembers Version=\"1.0\">");

            if (!string.IsNullOrEmpty(systemUserEmail))
            {
                string partyId = (systemUserId.HasValue && systemUserId.Value != Guid.Empty)
                    ? ToBracedUpper(systemUserId.Value)
                    : "";
                sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" />",
                    X(systemUserEmail), X(partyId));
            }

            if (isIncoming && !string.IsNullOrEmpty(fromSmtp)
                && !string.Equals(fromSmtp, systemUserEmail, StringComparison.OrdinalIgnoreCase))
            {
                sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"\" TypeCode=\"-1\" Name=\"\" />", X(fromSmtp));
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (recipients != null)
            {
                foreach (var r in recipients)
                {
                    var e = (r ?? "").Trim();
                    if (e.Length == 0) continue;
                    if (seen.Contains(e)) continue;
                    if (string.Equals(e, systemUserEmail, StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(e, fromSmtp, StringComparison.OrdinalIgnoreCase)) continue;

                    seen.Add(e);
                    sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"\" TypeCode=\"-1\" Name=\"\" />", X(e));
                }
            }

            sb.Append("</PartyMembers>");
            return sb.ToString();
        }

        private static string BuildCrmPartyInfoXmlForAppointment(string organizerSmtp, Guid? organizerSystemUserId)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("<PartyMembers Version=\"1.0\">");

            string partyId = (organizerSystemUserId.HasValue && organizerSystemUserId.Value != Guid.Empty)
                ? ToBracedUpper(organizerSystemUserId.Value)
                : "";
            string email = organizerSmtp ?? "";

            sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" ParticipationType=\"5\" />",
                X(email), X(partyId));
            sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" ParticipationType=\"7\" />",
                X(email), X(partyId));

            sb.Append("</PartyMembers>");
            return sb.ToString();
        }
    }
}
