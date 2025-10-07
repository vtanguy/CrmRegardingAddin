
// CrmMapiInterop.cs
// VS2015 + C#6 compatible
using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Security;


namespace CrmRegardingAddin
{
    public static class CrmMapiInterop
    {
        private const string DASL_BASE = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

        // UDF alignées sur ce que lit CrmLinkPane
        private const string P_LINKSTATE = "crmlinkstate";
        private const string P_CRMID       = "crmid"; // GUID CRM de l’activité (avec accolades)

        private const string P_SSS_TRACK = "crmSssPromoteTracker"; // on NE l’utilisera PAS pour les RDV
        private const string P_REGARDINGID = "crmregardingobjectid";
        private const string P_REGARDINGOT = "crmregardingobjecttypecode";
        private const string P_REGARDING = "crmregardingobject";
        private const string P_ORGID = "crmorgid";
        private const string P_PARTYINFO = "crmpartyinfo";
        private const string P_OWNER_SMTP = "crmownersmtp";
        private const string P_OWNER_SYSID = "crmownersystemuserid";

        private static string N(string name) { return DASL_BASE + name; }

        public class CrmPartyMember
        {
            public string Email { get; set; }
            public Guid? PartyId { get; set; }
            public int? TypeCode { get; set; }
            public string Name { get; set; }
        }

        private static string ToUnbracedLower(Guid id)
        {
            return id == Guid.Empty ? "" : id.ToString("D").ToLowerInvariant();
        }


        public static void ApplyMsCompatForMail(
            Outlook.MailItem mail,
            string regardingLogicalNameOrCode,
            Guid regardingId,
            string regardingDisplayName,
            string systemUserEmail,
            Guid? systemUserId,
            CrmPartyMember fromMember,
            IEnumerable<CrmPartyMember> recipients,
            bool isIncoming,
            Guid orgId
        )
        {
            if (mail == null) return;

            var pa = mail.PropertyAccessor;

            TrySet(pa, N(P_LINKSTATE), 2.0);
            // (disabled) TrySet(pa, N(P_SSS_TRACK), 0); // avoid SSS duplicate promotion for appointments
            if (regardingId != Guid.Empty)
                TrySet(pa, N(P_REGARDINGID), ToBracedUpper(regardingId));

            int apTypeCode; if (!string.IsNullOrEmpty(regardingLogicalNameOrCode) && int.TryParse(regardingLogicalNameOrCode, out apTypeCode)) TrySet(pa, N(P_REGARDINGOT), apTypeCode.ToString());

            if (!string.IsNullOrEmpty(regardingDisplayName))
                TrySet(pa, N(P_REGARDING), regardingDisplayName);

            if (orgId != Guid.Empty)
                TrySet(pa, N(P_ORGID), ToBracedUpper(orgId));

            string xml = BuildCrmPartyInfoXmlForMail(systemUserEmail, systemUserId, fromMember, recipients, isIncoming);
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
            string regardingLogicalNameOrCode,
            Guid orgId,
            string organizerSmtp,
            Guid? organizerSystemUserId)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;

            // 1) Indiquer le lien (le panneau lit 'crmlinkstate')
            TrySet(pa, N(P_LINKSTATE), 2.0);

            // 2) NE PAS bloquer SSS : surtout ne pas mettre crmSssPromoteTracker=0
            // TrySet(pa, N(P_SSS_TRACK), 0); // ❌ à NE PAS faire pour les rendez-vous

            // 3) Regarding
            if (regardingId != Guid.Empty)
                TrySet(pa, N(P_REGARDINGID), ToBracedUpper(regardingId));

            int typeCode;
            if (!string.IsNullOrEmpty(regardingLogicalNameOrCode) && int.TryParse(regardingLogicalNameOrCode, out typeCode))
                TrySet(pa, N(P_REGARDINGOT), typeCode.ToString());

            if (!string.IsNullOrEmpty(regardingDisplayName))
                TrySet(pa, N(P_REGARDING), regardingDisplayName);

            if (orgId != Guid.Empty)
                TrySet(pa, N(P_ORGID), ToBracedUpper(orgId));

            // 4) Infos d’organisateur attendues par le panneau (et utiles côté add-in MS)
            if (!string.IsNullOrEmpty(organizerSmtp))
                TrySet(pa, N(P_OWNER_SMTP), organizerSmtp);
            if (organizerSystemUserId.HasValue && organizerSystemUserId.Value != Guid.Empty)
                TrySet(pa, N(P_OWNER_SYSID), ToBracedUpper(organizerSystemUserId.Value));

            // 5) XML PartyInfo (organizer/required/optional selon ton implémentation)
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
            CrmPartyMember fromMember,
            IEnumerable<CrmPartyMember> recipients,
            bool isIncoming)
        {
            // Schéma attendu par l’addin Microsoft (vu dans tes dumps MFCMAPI) :
            // <PartyMembers Version="1.0">
            //   <Member Email="..." PartyId="optional" TypeCode="8 or -1" Name="..." />
            //   ...
            // </PartyMembers>
            //
            // - systemuser => TypeCode="8", PartyId = GUID sans accolades, en minuscules
            // - autres adresses => TypeCode renseigné si connu (contact=2, account=1, lead=4...), sinon -1
            // - Name doit être non vide (fallback = email)

            var sb = new StringBuilder();
            sb.Append("<PartyMembers Version=\"1.0\">");

            Action<CrmPartyMember> writeMember = member =>
            {
                if (member == null) return;

                var email = (member.Email ?? "").Trim();
                if (email.Length == 0) return;

                int typeCode = member.TypeCode ?? -1;
                Guid? partyId = member.PartyId;

                var safeName = string.IsNullOrWhiteSpace(member.Name) ? email : member.Name;
                string partyIdAttr = "";
                if (partyId.HasValue && partyId.Value != Guid.Empty)
                {
                    // IMPORTANT: pas d’accolades, en minuscules
                    partyIdAttr = $" PartyId=\"{ToUnbracedLower(partyId.Value)}\"";
                }

                sb.Append("<Member");
                sb.Append($" Email=\"{SecurityElement.Escape(email)}\"");
                sb.Append(partyIdAttr);
                sb.Append($" TypeCode=\"{typeCode}\"");
                sb.Append($" Name=\"{SecurityElement.Escape(safeName)}\"");
                sb.Append(" />");
            };

            // 1) FROM / TO selon entrant/sortant
            if (isIncoming)
            {
                // Entrant: FROM = expéditeur SMTP
                writeMember(fromMember);

                // TO = systemuser (TypeCode=8, avec PartyId si dispo)
                var systemUserMember = BuildSystemUserMember(systemUserEmail, systemUserId);
                writeMember(systemUserMember);
            }
            else
            {
                // Sortant: FROM = systemuser (TypeCode=8)
                var systemUserMember = BuildSystemUserMember(systemUserEmail, systemUserId);
                writeMember(systemUserMember);

                // TO = destinataires (TypeCode=-1, pas de PartyId)
                if (recipients != null)
                {
                    foreach (var r in recipients)
                    {
                        writeMember(r);
                    }
                }
            }

            sb.Append("</PartyMembers>");
            return sb.ToString();
        }

        private static CrmPartyMember BuildSystemUserMember(string email, Guid? systemUserId)
        {
            if (string.IsNullOrWhiteSpace(email)) return null;
            return new CrmPartyMember
            {
                Email = email,
                Name = email,
                PartyId = systemUserId,
                TypeCode = 8
            };
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

    

    public static void TagAfterCrmAppointmentCreate(
    Outlook.AppointmentItem appt,
    Guid crmApptId,
    Guid regardingId,
    string regardingDisplayName,
    string regardingLogicalNameOrCode,
    Guid orgId,
    string ownerSmtp,
    Guid? ownerSystemUserId)
{
    if (appt == null) return;
    var pa = appt.PropertyAccessor;

    TrySet(pa, N(P_LINKSTATE), 2.0);

    if (crmApptId != Guid.Empty)
        TrySet(pa, N(P_CRMID), ToBracedUpper(crmApptId));

    if (regardingId != Guid.Empty)
        TrySet(pa, N(P_REGARDINGID), ToBracedUpper(regardingId));

    int typeCode;
    if (!string.IsNullOrEmpty(regardingLogicalNameOrCode) && int.TryParse(regardingLogicalNameOrCode, out typeCode))
        TrySet(pa, N(P_REGARDINGOT), typeCode.ToString());

    if (!string.IsNullOrEmpty(regardingDisplayName))
        TrySet(pa, N(P_REGARDING), regardingDisplayName);

    if (orgId != Guid.Empty)
        TrySet(pa, N(P_ORGID), ToBracedUpper(orgId));

    if (!string.IsNullOrEmpty(ownerSmtp))
        TrySet(pa, N(P_OWNER_SMTP), ownerSmtp);
    if (ownerSystemUserId.HasValue && ownerSystemUserId.Value != Guid.Empty)
        TrySet(pa, N(P_OWNER_SYSID), ToBracedUpper(ownerSystemUserId.Value));

    try { appt.Save(); } catch { }
}
}


}
