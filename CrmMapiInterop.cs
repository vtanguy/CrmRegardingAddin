// CrmMapiInterop.cs
// VS2015 + C#6 compatible
using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    /// <summary>
    /// Ecrit / supprime les propriétés MAPI nommées comme l’addin Microsoft Dynamics CRM
    /// afin que nos liens soient visibles/compatibles côté addin Microsoft.
    /// </summary>
    public static class CrmMapiInterop
    {
        // Base des propriétés nommées PS_PUBLIC_STRINGS
        private const string DASL_BASE = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

        // Noms de propriétés observés dans tes exports MFCMAPI
        private const string P_LINKSTATE = "crmLinkState";           // double -> 2.0
        private const string P_SSS_TRACK = "crmSssPromoteTracker";   // int -> 0
        private const string P_REGARDINGID = "crmRegardingId";         // string -> "{GUID}"
        private const string P_REGARDINGOT = "crmRegardingObjectType"; // string -> "2" (ex.)
        private const string P_REGARDING = "Regarding";              // string -> libellé
        private const string P_ORGID = "crmorgid";               // string -> "{ORG_GUID}"
        private const string P_PARTYINFO = "crmpartyinfo";           // string (XML)

        private static string N(string name) { return DASL_BASE + name; }

        // ---------------- MAIL ----------------

        /// <summary>
        /// Pose les propriétés MAPI "façon Microsoft" sur un MailItem.
        /// </summary>
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
            Guid orgId // <— paramètre ajouté
        )
        {
            if (mail == null) return;

            var pa = mail.PropertyAccessor;

            // 1) Etats / identifiants
            try { pa.SetProperty(N(P_LINKSTATE), 2.0); } catch { }
            try { pa.SetProperty(N(P_SSS_TRACK), 0); } catch { }

            if (regardingId != Guid.Empty)
            {
                try { pa.SetProperty(N(P_REGARDINGID), ToBracedUpper(regardingId)); } catch { }
            }

            // Si tu disposes d’un code d’objet (ex. "2" pour contact), mets-le ; sinon laisse tel quel.
            int typeCode;
            if (!string.IsNullOrEmpty(regardingLogicalNameOrCode) && int.TryParse(regardingLogicalNameOrCode, out typeCode))
            {
                try { pa.SetProperty(N(P_REGARDINGOT), typeCode.ToString()); } catch { }
            }

            if (!string.IsNullOrEmpty(regardingDisplayName))
            {
                try { pa.SetProperty(N(P_REGARDING), regardingDisplayName); } catch { }
            }

            if (orgId != Guid.Empty)
            {
                try { pa.SetProperty(N(P_ORGID), ToBracedUpper(orgId)); } catch { }
            }

            // 2) crmpartyinfo (XML)
            try
            {
                string xml = BuildCrmPartyInfoXmlForMail(systemUserEmail, systemUserId, fromSmtp, recipients, isIncoming);
                if (!string.IsNullOrEmpty(xml))
                    pa.SetProperty(N(P_PARTYINFO), xml);
            }
            catch { }

            try { mail.Save(); } catch { }
        }

        /// <summary>
        /// Supprime les propriétés MAPI "suivi" sur un MailItem.
        /// </summary>
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

        // ---------------- RENDEZ-VOUS ----------------

        /// <summary>
        /// Pose les propriétés MAPI "façon Microsoft" sur un AppointmentItem.
        /// </summary>
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

            try { pa.SetProperty(N(P_LINKSTATE), 2.0); } catch { }
            try { pa.SetProperty(N(P_SSS_TRACK), 0); } catch { }

            if (regardingId != Guid.Empty)
            {
                try { pa.SetProperty(N(P_REGARDINGID), ToBracedUpper(regardingId)); } catch { }
            }

            if (!string.IsNullOrEmpty(regardingDisplayName))
            {
                try { pa.SetProperty(N(P_REGARDING), regardingDisplayName); } catch { }
            }

            // crmpartyinfo pour rendez-vous : organizer en double (ParticipationType 5 et 7)
            try
            {
                string xml = BuildCrmPartyInfoXmlForAppointment(organizerSmtp, organizerSystemUserId);
                if (!string.IsNullOrEmpty(xml))
                    pa.SetProperty(N(P_PARTYINFO), xml);
            }
            catch { }

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

        // ---------------- Helpers ----------------

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
            var sb = new StringBuilder();
            sb.Append("<PartyMembers Version=\"1.0\">");

            // 1) System user (TypeCode 8)
            if (!string.IsNullOrEmpty(systemUserEmail))
            {
                string partyId = (systemUserId.HasValue && systemUserId.Value != Guid.Empty)
                    ? ToBracedUpper(systemUserId.Value)
                    : "";
                sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" />",
                    X(systemUserEmail), X(partyId));
            }

            // 2) Expéditeur (TypeCode -1 si différent du user et si entrant)
            if (isIncoming && !string.IsNullOrEmpty(fromSmtp)
                && !string.Equals(fromSmtp, systemUserEmail, StringComparison.OrdinalIgnoreCase))
            {
                sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"\" TypeCode=\"-1\" Name=\"\" />", X(fromSmtp));
            }

            // 3) Destinataires To/Cc/Bcc (TypeCode -1) sans doublons et en excluant system user / sender déjà ajoutés
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
            var sb = new StringBuilder();
            sb.Append("<PartyMembers Version=\"1.0\">");

            string partyId = (organizerSystemUserId.HasValue && organizerSystemUserId.Value != Guid.Empty)
                ? ToBracedUpper(organizerSystemUserId.Value)
                : "";
            string email = organizerSmtp ?? "";

            // Organizer avec ParticipationType 5 et 7 (constaté dans tes dumps)
            sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" ParticipationType=\"5\" />",
                X(email), X(partyId));
            sb.AppendFormat("<Member Email=\"{0}\" PartyId=\"{1}\" TypeCode=\"8\" Name=\"\" ParticipationType=\"7\" />",
                X(email), X(partyId));

            sb.Append("</PartyMembers>");
            return sb.ToString();
        }
    }
}
