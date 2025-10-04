// CrmMapiInterop.cs
// VS2015 + C#6 compatible
using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    /// <summary>
    /// Writes the same named MAPI properties as the official Dynamics CRM for Outlook add-in,
    /// so both add-ins recognize links ("Suivi") the same way.
    /// </summary>
    public static class CrmMapiInterop
    {
        // PS_PUBLIC_STRINGS named-property base
        private const string DASL_BASE = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

        // Suffixes by type
        private const string PT_UNICODE = "/0x0000001F";
        private const string PT_INT32   = "/0x00000003";
        private const string PT_DOUBLE  = "/0x00000005";

        // Known Dynamics named props (see MFCMAPI dumps you provided)
        private const string P_CRM_LINKSTATE         = "crmLinkState";
        private const string P_CRM_REGARDING_ID      = "crmRegardingId";
        private const string P_CRM_REGARDING_TYPE    = "crmRegardingObjectType";
        private const string P_CRM_SSS_TRACKER       = "crmSssPromoteTracker";
        private const string P_CRM_PARTYINFO         = "crmpartyinfo";
        private const string P_REGARDING_DISPLAY     = "Regarding";
        private const string P_CRM_OWNERID           = "crmOwnerId"; // optional (appointments)

        private static string D(string name, string suffix) { return DASL_BASE + name + suffix; }

        /// <summary>
        /// Apply MS-compat flags on a MailItem after (or while) creating the CRM email & regarding.
        /// - regardingLogicalName: e.g. "account" (not strictly used by MS add-in but we keep it)
        /// - regardingId: the GUID of the CRM record (no braces or with braces both accepted; we will normalize with braces)
        /// - systemUserEmail: the Outlook/CRM user primary SMTP (To party for incoming; From party for outgoing)
        /// - systemUserId: the CRM GUID of the current user (optional for partyinfo)
        /// - fromSmtp: actual sender SMTP (for incoming) or current user (for outgoing)
        /// - recipients: list of SMTP addresses that were To/CC/Bcc (used to enrich crmpartyinfo)
        /// </summary>
        public static void ApplyMsCompatForMail(
            Outlook.MailItem mail,
            string regardingLogicalName,
            Guid regardingId,
            string regardingDisplayName,
            string systemUserEmail,
            Guid? systemUserId,
            string fromSmtp,
            IEnumerable<string> recipients,
            bool isIncoming)
        {
            if (mail == null) return;

            var pa = mail.PropertyAccessor;

            // Normalize GUID with braces like {XXXXXXXX-...}
            var regIdBraced = ToBracedGuid(regardingId);

            // 1) Required trio: RegardingId, RegardingObjectType (string), Regarding (display)
            pa.SetProperty(D(P_CRM_REGARDING_ID,   PT_UNICODE), regIdBraced);
            int objectTypeCode;
            if (int.TryParse(regardingLogicalName, out objectTypeCode))
            {
                pa.SetProperty(D(P_CRM_REGARDING_TYPE, PT_UNICODE), objectTypeCode.ToString());
            }
            if (!string.IsNullOrEmpty(regardingDisplayName))
                pa.SetProperty(D(P_REGARDING_DISPLAY, PT_UNICODE), regardingDisplayName);

            // 2) Link state flags
            pa.SetProperty(D(P_CRM_LINKSTATE, PT_DOUBLE), 2.0);
            pa.SetProperty(D(P_CRM_SSS_TRACKER, PT_INT32), 0);

            // 3) PartyInfo XML
            var xml = BuildCrmPartyInfoXmlForMail(systemUserEmail, systemUserId, fromSmtp, recipients, isIncoming);
            if (!string.IsNullOrEmpty(xml))
            {
                pa.SetProperty(D(P_CRM_PARTYINFO, PT_UNICODE), xml);
            }

            mail.Save();
        }

        /// <summary>
        /// Remove MS-compat properties from a MailItem (unlink only — does not delete CRM record).
        /// </summary>
        public static void RemoveMsCompatFromMail(Outlook.MailItem mail)
        {
            if (mail == null) return;
            var pa = mail.PropertyAccessor;
            TryDelete(pa, D(P_CRM_LINKSTATE,     PT_DOUBLE));
            TryDelete(pa, D(P_CRM_SSS_TRACKER,   PT_INT32));
            TryDelete(pa, D(P_CRM_REGARDING_ID,  PT_UNICODE));
            TryDelete(pa, D(P_CRM_REGARDING_TYPE,PT_UNICODE));
            TryDelete(pa, D(P_CRM_PARTYINFO,     PT_UNICODE));
            TryDelete(pa, D(P_REGARDING_DISPLAY, PT_UNICODE));
            mail.Save();
        }

        /// <summary>
        /// Apply MS-compat flags on an AppointmentItem (rendez-vous).
        /// </summary>
        public static void ApplyMsCompatForAppointment(
            Outlook.AppointmentItem appt,
            Guid regardingId,
            string regardingDisplayName,
            string organizerSmtp,
            Guid? organizerSystemUserId)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;

            var regIdBraced = ToBracedGuid(regardingId);
            pa.SetProperty(D(P_CRM_REGARDING_ID,   PT_UNICODE), regIdBraced);
            if (!string.IsNullOrEmpty(regardingDisplayName))
                pa.SetProperty(D(P_REGARDING_DISPLAY, PT_UNICODE), regardingDisplayName);
            pa.SetProperty(D(P_CRM_LINKSTATE, PT_DOUBLE), 2.0);
            pa.SetProperty(D(P_CRM_SSS_TRACKER, PT_INT32), 0);

            // PartyInfo: organizer with ParticipationType 5 & 7
            var xml = BuildCrmPartyInfoXmlForAppointment(organizerSmtp, organizerSystemUserId);
            if (!string.IsNullOrEmpty(xml))
                pa.SetProperty(D(P_CRM_PARTYINFO, PT_UNICODE), xml);

            appt.Save();
        }

        public static void RemoveMsCompatFromAppointment(Outlook.AppointmentItem appt)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;
            TryDelete(pa, D(P_CRM_LINKSTATE,     PT_DOUBLE));
            TryDelete(pa, D(P_CRM_SSS_TRACKER,   PT_INT32));
            TryDelete(pa, D(P_CRM_REGARDING_ID,  PT_UNICODE));
            TryDelete(pa, D(P_CRM_REGARDING_TYPE,PT_UNICODE));
            TryDelete(pa, D(P_CRM_PARTYINFO,     PT_UNICODE));
            TryDelete(pa, D(P_REGARDING_DISPLAY, PT_UNICODE));
            appt.Save();
        }

        // ------------------ helpers ------------------

        private static void TryDelete(Outlook.PropertyAccessor pa, string dasl)
        {
            try { pa.DeleteProperty(dasl); } catch { /* ignore */ }
        }

        private static string ToBracedGuid(Guid id)
        {
            return "{" + id.ToString().ToUpper() + "}";
        }

        private static string XmlEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Replace("&", "&amp;").
                     Replace("\"", "&quot;").
                     Replace("<", "&lt;").
                     Replace(">", "&gt;");
        }

        private static string BuildCrmPartyInfoXmlForMail(
            string systemUserEmail,
            Guid? systemUserId,
            string fromSmtp,
            IEnumerable<string> recipients,
            bool isIncoming)
        {
            var members = new List<string>();

            var sysPartyId = systemUserId.HasValue ? "{" + systemUserId.Value.ToString().ToUpper() + "}" : "";
            var sysEmail = XmlEscape(systemUserEmail ?? string.Empty);
            var senderEmail = XmlEscape(fromSmtp ?? string.Empty);

            if (!string.IsNullOrEmpty(sysEmail))
                members.Add(string.Format("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" />",
                    sysEmail, XmlEscape(sysPartyId)));

            if (isIncoming && !string.IsNullOrEmpty(senderEmail))
                members.Add(string.Format("<Member Email=\"{0}\" PartyId=\"\" TypeCode=\"-1\" Name=\"\" />", senderEmail));

            if (recipients != null)
            {
                foreach (var r in recipients)
                {
                    var e = (r ?? string.Empty).Trim();
                    if (e.Length == 0) continue;
                    if (string.Equals(e, systemUserEmail, StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(e, fromSmtp, StringComparison.OrdinalIgnoreCase)) continue;
                    members.Add(string.Format("<Member Email=\"{0}\" PartyId=\"\" TypeCode=\"-1\" Name=\"\" />", XmlEscape(e)));
                }
            }

            if (members.Count == 0) return null;
            return "<PartyMembers Version=\"1.0\">" + string.Join("", members.ToArray()) + "</PartyMembers>";
        }

        private static string BuildCrmPartyInfoXmlForAppointment(string organizerSmtp, Guid? organizerSystemUserId)
        {
            var orgEmail = XmlEscape(organizerSmtp ?? string.Empty);
            var orgId = organizerSystemUserId.HasValue ? "{" + organizerSystemUserId.Value.ToString().ToUpper() + "}" : "";

            var member1 = string.Format("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" ParticipationType=\"5\" />", orgEmail, XmlEscape(orgId));
            var member2 = string.Format("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" ParticipationType=\"7\" />", orgEmail, XmlEscape(orgId));
            return "<PartyMembers Version=\"1.0\">" + member1 + member2 + "</PartyMembers>";
        }
    }
}
