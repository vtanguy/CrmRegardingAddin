using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Security;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;

namespace CrmRegardingAddin
{
    public static class LinkingApi
    {
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";
        private const int OTC_SYSTEMUSER = 8;
        private const int OTC_ACCOUNT = 1;
        private const int OTC_CONTACT = 2;
        private const int OTC_LEAD = 4;
        private const int OTC_APPOINTMENT = 4201;
        private const int OTC_EMAIL = 4202;


        // === MS add-in string-named aliases ===
        private const string DASL_RegId_String_Lower   = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmregardingobjectid";
        private const string DASL_RegId_String_Camel   = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmRegardingObjectId";
        private const string DASL_RegId_String_Old     = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmRegardingId";
        private const string DASL_RegTypeCode_Lower    = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmregardingobjecttypecode";
        private const string DASL_RegTypeCode_Camel    = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmRegardingObjectTypeCode";

        private static class DaslId
        {
            public const string UserProperties     = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C8"; // PT_BINARY
            public const string LinkState            = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80BD"; // PT_DOUBLE
            public const string RegardingId          = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C9"; // PT_UNICODE
            public const string RegardingObjectType  = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80CA"; // PT_UNICODE
            public const string EntryID              = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C3"; // PT_UNICODE
            public const string CrmId                = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C4"; // PT_UNICODE
            public const string OrgId                = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C5"; // PT_UNICODE
            public const string CrmXml               = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80CF"; // PT_UNICODE
            public const string ObjectTypeCode       = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80D1"; // PT_DOUBLE
            public const string PartyInfo            = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80D5"; // PT_UNICODE
            public const string AsyncSend            = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80D7"; // PT_DOUBLE
            public const string CrmMessageId         = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80D8"; // PT_UNICODE
            public const string SssPromoteTracker    = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80DE"; // PT_LONG
            public const string TrackedBySender      = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80DF"; // PT_BOOLEAN
        }

        private static class DaslStr // string-named props (MS add-in compatibility)
        {
            public const string LinkState                   = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmLinkState";
            public const string RegardingId_Lower           = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingobjectid";
            public const string RegardingId_Camel           = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectId";
            public const string RegardingId_Old             = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingId"; // 0x80F3
            public const string RegardingTypeCode_Lower     = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingobjecttypecode";
            public const string RegardingTypeCode_Camel     = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectTypeCode";
            public const string RegardingType_Old           = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectType"; // 0x80F2

            public const string RegardingLabel              = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/Regarding";

            public const string EntryID                     = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmEntryID";
            public const string CrmId                       = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmid";
            public const string OrgId                       = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmorgid";
            public const string CrmXml                      = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmxml";
            public const string ObjectTypeCode              = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmObjectTypeCode";
            public const string PartyInfo                   = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmpartyinfo";
            public const string AsyncSend                   = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmAsyncSend";
            public const string CrmMessageId                = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmmessageid";
            public const string SssPromoteTracker           = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmSssPromoteTracker";
            public const string TrackedBySender             = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmTrackedBySender";

            // Owner fields observed in OK CSV
            public const string OwnerId                     = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmOwnerId";      // 0x8524
            public const string OwnerIdType                 = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmOwnerIdType";  // 0x8525
        }

        private static string Braced(Guid g) => "{" + g.ToString().ToUpperInvariant() + "}";

        private static void TrySet(Outlook.PropertyAccessor pa, string dasl, object value) { try { pa.SetProperty(dasl, value); } catch { } }
        private static string TryGet(Outlook.PropertyAccessor pa, params string[] dasls)
        {
            foreach (var d in dasls)
            {
                try
                {
                    var o = pa.GetProperty(d);
                    if (o != null)
                    {
                        var s = Convert.ToString(o, CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                    }
                } catch { }
            }
            return null;
        }
        private static int? TryGetInt(Outlook.PropertyAccessor pa, params string[] dasls)
        {
            var s = TryGet(pa, dasls);
            if (string.IsNullOrWhiteSpace(s)) return null;
            int v; return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? (int?)v : null;
        }

        private static int? ResolveTypeCode(IOrganizationService org, string logicalName)
        {
            if (org == null || string.IsNullOrWhiteSpace(logicalName)) return null;
            try
            {
                var req = new RetrieveEntityRequest { LogicalName = logicalName, EntityFilters = EntityFilters.Entity };
                var resp = (RetrieveEntityResponse)org.Execute(req);
                return resp.EntityMetadata.ObjectTypeCode;
            }
            catch { return null; }
        }
        private static string LogicalNameFromTypeCode(IOrganizationService org, int typeCode)
        {
            switch (typeCode)
            {
                case 1: return "account"; case 2: return "contact"; case 3: return "opportunity"; case 4: return "lead";
                case 8: return "systemuser"; case 4201: return "appointment"; case 4202: return "email";
            }
            try
            {
                var allReq = new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Entity, RetrieveAsIfPublished = true };
                var allResp = (RetrieveAllEntitiesResponse)org.Execute(allReq);
                foreach (var md in allResp.EntityMetadata)
                    if (md.ObjectTypeCode.HasValue && md.ObjectTypeCode.Value == typeCode) return md.LogicalName;
            }
            catch { }
            return null;
        }
        private static Guid? GetOrgId(IOrganizationService org)
        {
            try { var w = (WhoAmIResponse)org.Execute(new WhoAmIRequest()); return w?.OrganizationId; } catch { return null; }
        }

        // ======================== PREPARE ========================

        public static void PrepareAppointmentLinkInOutlookStore(IOrganizationService org, Outlook.AppointmentItem appt, Guid regardingId, string regardingLogicalName, string regardingReadableName)
        {
            if (appt == null) return;
            int otc = (ResolveTypeCode(org, regardingLogicalName) ?? 0);
            // MS add-in compatible (string + id variants)
            CrmMapiInterop.SetCrmLinkPropsForAppointment(appt, regardingId, otc, null, null, null);

            var pa = appt.PropertyAccessor;
            // Explicitly write the exact props observed in the "OK" CSV so the MS add-in resolves Regarding:
            // - 0x80F3 crmRegardingId (braced GUID)
            // - 0x80F2 crmRegardingObjectType (numeric as string)
            // - crmOwnerId / crmOwnerIdType (owner = current user, type 8)
            // - crmObjectTypeCode (4201) as DOUBLE + string for completeness
            try
            {
                string reg = Braced(regardingId);
                if (!string.IsNullOrWhiteSpace(reg))
                {
                    TrySet(pa, DaslStr.RegardingId_Old, reg);            // 0x80F3
                    TrySet(pa, DaslStr.RegardingId_Lower, reg);
                    TrySet(pa, DaslStr.RegardingId_Camel, reg);
                }

                if (otc > 0)
                {
                    string otcStr = otc.ToString(CultureInfo.InvariantCulture);
                    TrySet(pa, DaslStr.RegardingType_Old, otcStr);        // 0x80F2 expected by MS add-in
                    TrySet(pa, DaslStr.RegardingTypeCode_Lower, otcStr);
                    TrySet(pa, DaslStr.RegardingTypeCode_Camel, otcStr);
                }

                // Owner (WhoAmI) to mirror OK CSV
                try
                {
                    var who = (WhoAmIResponse)org.Execute(new WhoAmIRequest());
                    if (who != null && who.UserId != Guid.Empty)
                    {
                        TrySet(pa, DaslStr.OwnerId, Braced(who.UserId));          // 0x8524
                        TrySet(pa, DaslStr.OwnerIdType, OTC_SYSTEMUSER.ToString(CultureInfo.InvariantCulture)); // 0x8525
                    }
                } catch { }

                // crmObjectTypeCode of the item (appointment) – DOUBLE + string
                TrySet(pa, DaslId.ObjectTypeCode, (double)4201);
                TrySet(pa, DaslStr.ObjectTypeCode, "4201");

                // LinkState as DOUBLE on the string-named prop too
                TrySet(pa, DaslId.LinkState, 1.0);   // numeric id
                TrySet(pa, DaslStr.LinkState, "1");  // ensure 0x80B4 is PT_DOUBLE, not PT_UNICODE
            }
            catch { }

            // Write crmpartyinfo (PartyMembers XML) for appointments
            try {
                var __xml = BuildPartyInfo(org, appt);
                if (!string.IsNullOrWhiteSpace(__xml)) { TrySet(pa, DaslStr.PartyInfo, __xml); TrySet(pa, DaslId.PartyInfo, __xml); try { appt.Save(); } catch { } }
            } catch { }

            try { EnsureUserPropertiesForAppointment(appt); } catch { }
            try { appt.Save(); } catch { }
        }

        public static void PrepareMailLinkInOutlookStore(IOrganizationService org, Outlook.MailItem mi, Guid regardingId, string regardingLogicalName, string regardingReadableName)
        {
            if (mi == null) return; var pa = mi.PropertyAccessor;

            // Add MS add-in numeric regarding typecode aliases for mails (compat)
            try
            {
                int __otcMail = 0;
                try
                {
                    var __name = regardingLogicalName;
                    if (!string.IsNullOrWhiteSpace(__name))
                    {
                        var __tmp = ResolveTypeCode(org, __name);
                        if (__tmp.HasValue) __otcMail = __tmp.Value;
                    }
                } catch { }
                if (__otcMail > 0)
                {
                    var __otcStr = __otcMail.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    TrySet(pa, DaslStr.RegardingTypeCode_Lower, __otcStr);
                    TrySet(pa, DaslStr.RegardingTypeCode_Camel, __otcStr);
                    TrySet(pa, DaslStr.RegardingType_Old, __otcStr); // keep 0x80F2 synced too
                }
            }
            catch { }

            // crmRegardingId on all aliases
            TrySet(pa, DaslId.LinkState, 1.0);                         TrySet(pa, DaslStr.LinkState, "1");
            TrySet(pa, DaslId.RegardingId, Braced(regardingId));       TrySet(pa, DaslStr.RegardingId_Camel, Braced(regardingId)); TrySet(pa, DaslStr.RegardingId_Lower, Braced(regardingId)); TrySet(pa, DaslStr.RegardingId_Old, Braced(regardingId));
            TrySet(pa, DaslId.RegardingObjectType, regardingLogicalName); TrySet(pa, DaslStr.RegardingType_Old, TryGet(pa, DaslStr.RegardingType_Old)); // keep whatever we set above
            if (!string.IsNullOrWhiteSpace(regardingReadableName)) { TrySet(pa, DaslStr.RegardingLabel, regardingReadableName); }
            TrySet(pa, DaslId.SssPromoteTracker, 1);                   TrySet(pa, DaslStr.SssPromoteTracker, "1");

            // NEW: Write crmpartyinfo (PartyMembers XML) for mails (To, Cc, Bcc, From)
            try
            {
                var xml = BuildPartyInfo(org, mi);
                if (!string.IsNullOrWhiteSpace(xml))
                {
                    TrySet(pa, DaslStr.PartyInfo, xml);
                    TrySet(pa, DaslId.PartyInfo, xml);
                }
                try { EnsureUserPropertiesForMail(mi); } catch { }
            } catch { }

            try { EnsureUserPropertiesForMail(mi); } catch { }
            try { mi.Save(); } catch { }
        }
        // ======================== PREPARED STATE HELPERS ========================

        /// <summary>
        /// True si l’email a été préparé (crmLinkState >= 1) ET qu’un RegardingId est présent.
        /// </summary>
        public static bool HasPreparedMailLink(Outlook.MailItem mi)
        {
            if (mi == null) return false;
            try
            {
                var pa = mi.PropertyAccessor;
                if (pa == null) return false;

                // LinkState peut être stocké en DOUBLE (id) ou en string (alias)
                string ls = TryGet(pa, DaslStr.LinkState, DaslId.LinkState);
                double d;
                if (!string.IsNullOrWhiteSpace(ls)
                    && double.TryParse(ls, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out d))
                {
                    if (d < 1.0) return false;
                }
                else
                {
                    try
                    {
                        object o = pa.GetProperty(DaslId.LinkState);
                        if (o is double && ((double)o) < 1.0) return false;
                    }
                    catch { }
                }

                var reg = TryGet(pa,
                                 DaslId.RegardingId,
                                 DaslStr.RegardingId_Lower,
                                 DaslStr.RegardingId_Camel,
                                 DaslStr.RegardingId_Old);
                return !string.IsNullOrWhiteSpace(reg);
            }
            catch { return false; }
        }

        /// <summary>
        /// Re-construit crmpartyinfo (MAIL) avec les destinataires finaux,
        /// assure la présence de 0x80C8 (UserProperties), puis commit CRM + finalize Outlook.
        /// </summary>
        public static Guid FinalizePreparedMailIfPossible(IOrganizationService org, Outlook.MailItem mi)
        {
            if (org == null || mi == null) throw new System.ArgumentNullException();
            try
            {
                var pa = mi.PropertyAccessor;

                // Rebuild crmpartyinfo (version MAIL MS-compatible que tu as déjà)
                var xml = BuildPartyInfo(org, mi);
                if (!string.IsNullOrWhiteSpace(xml))
                {
                    TrySet(pa, DaslStr.PartyInfo, xml);
                    TrySet(pa, DaslId.PartyInfo, xml);
                }

                // Forcer la (re)génération du BLOB 0x80C8
                EnsureUserPropertiesForMail(mi);
                try { mi.Save(); } catch { }
            }
            catch { /* non bloquant avant commit CRM */ }

            // Commit + finalize
            var id = CommitMailLinkToCrm(org, mi);
            FinalizeMailLinkInOutlookStoreAfterCrmCommit(org, mi, id);
            return id;
        }

        // ======================== COMMIT (CRM) ========================

        // APPOINTMENT
        public static Guid CommitAppointmentLinkToCrm(IOrganizationService org, Outlook.AppointmentItem appt)
        {
            if (org == null || appt == null) throw new ArgumentNullException();

            var pa = appt.PropertyAccessor;
            var regId  = TryGet(pa, DaslId.RegardingId, DaslStr.RegardingId_Lower, DaslStr.RegardingId_Camel, DaslStr.RegardingId_Old);
            var regOTs = TryGet(pa, DaslId.RegardingObjectType, DaslStr.RegardingTypeCode_Lower, DaslStr.RegardingTypeCode_Camel, DaslStr.RegardingType_Old);
            if (string.IsNullOrWhiteSpace(regId) || string.IsNullOrWhiteSpace(regOTs))
                throw new InvalidOperationException("Propriétés de lien insuffisantes sur le rendez-vous Outlook.");

            Guid regardingId; Guid.TryParse(regId.Trim('{','}'), out regardingId);
            int typeCode;
            if (!int.TryParse(regOTs, NumberStyles.Integer, CultureInfo.InvariantCulture, out typeCode))
            {
                var otc = ResolveTypeCode(org, regOTs);
                if (!otc.HasValue) throw new InvalidOperationException("Type de l'enregistrement CRM non déterminé.");
                typeCode = otc.Value;
            }
            var logicalName = LogicalNameFromTypeCode(org, typeCode) ?? "account";

            string globalId = null; try { globalId = appt.GlobalAppointmentID; } catch { }
            if (string.IsNullOrWhiteSpace(globalId))
                throw new InvalidOperationException("Impossible de récupérer le GlobalAppointmentID.");

            var existingId = FindCrmAppointmentIdByGlobalId(org, globalId);
            if (existingId != Guid.Empty)
            {
                var upd = new Entity("appointment") { Id = existingId };
                upd["regardingobjectid"] = new EntityReference(logicalName, regardingId);
                TryBackfillApptPartiesIfEmpty(org, existingId, upd, appt);
                org.Update(upd);
                return existingId;
            }

            var e = new Entity("appointment");
            e["subject"] = Safe(() => appt.Subject) ?? "";
            try { e["description"] = Safe(() => appt.Body) ?? ""; } catch { }
            try { e["scheduledstart"] = appt.Start; } catch { }
            try { e["scheduledend"] = appt.End; } catch { }
            e["regardingobjectid"] = new EntityReference(logicalName, regardingId);
            e["globalobjectid"] = globalId;
            e["organizer"] = ActivityPartyBuilder.BuildOrganizer(org);
            e["requiredattendees"] = ActivityPartyBuilder.BuildApptRecipientsFromRecipients(org, appt, true);
            e["optionalattendees"] = ActivityPartyBuilder.BuildApptRecipientsFromRecipients(org, appt, false);

            var id = org.Create(e);
            return id;
        }

        // MAIL
        public static Guid CommitMailLinkToCrm(IOrganizationService org, Outlook.MailItem mi)
        {
            if (org == null || mi == null) throw new ArgumentNullException();
            var pa = mi.PropertyAccessor;

            var regardingIdStr = TryGet(pa, DaslId.RegardingId, DaslStr.RegardingId_Camel, DaslStr.RegardingId_Lower, DaslStr.RegardingId_Old);
            var regardingLogical = TryGet(pa, DaslId.RegardingObjectType, DaslStr.RegardingType_Old);
            if (string.IsNullOrWhiteSpace(regardingIdStr) || string.IsNullOrWhiteSpace(regardingLogical))
                throw new InvalidOperationException("Propriétés de lien insuffisantes sur l'élément Outlook.");

            Guid regardingId; Guid.TryParse(regardingIdStr.Trim('{', '}'), out regardingId);

            var e = new Entity("email");
            e["subject"] = mi.Subject ?? "";
            try { e["description"] = mi.HTMLBody ?? mi.Body ?? ""; } catch { e["description"] = mi.Body ?? ""; }
            
            // Map "regardingLogical" if it contains a numeric type code (e.g., "2" for contact).
            string __regardingLogical = regardingLogical;
            int __regardingTypeCode;
            if (int.TryParse(regardingLogical, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out __regardingTypeCode))
            {
                var __ln = LogicalNameFromTypeCode(org, __regardingTypeCode);
                if (!string.IsNullOrWhiteSpace(__ln)) __regardingLogical = __ln;
            }
    
            e["regardingobjectid"] = new EntityReference(__regardingLogical, regardingId);

            var senderSmtp = ActivityPartyBuilder.GetSenderSmtp(mi);
            var currentUserEmail = ActivityPartyBuilder.GetCurrentUserEmail(org);
            bool isOutgoing = (senderSmtp != null && currentUserEmail != null && string.Equals(senderSmtp, currentUserEmail, StringComparison.OrdinalIgnoreCase));
            e["directioncode"] = isOutgoing;
            e["from"] = ActivityPartyBuilder.BuildFromForMail(org, mi);
            e["to"] = ActivityPartyBuilder.BuildRecipientsFromOutlook(org, mi, Outlook.OlMailRecipientType.olTo);
            e["cc"] = ActivityPartyBuilder.BuildRecipientsFromOutlook(org, mi, Outlook.OlMailRecipientType.olCC);
            e["bcc"] = ActivityPartyBuilder.BuildRecipientsFromOutlook(org, mi, Outlook.OlMailRecipientType.olBCC);

            try { var msgId = MailUtil.GetInternetMessageId(mi); if (!string.IsNullOrWhiteSpace(msgId)) e["messageid"] = msgId; } catch { }

            var id = org.Create(e);

            // --- Upload Outlook attachments to CRM (activitymimeattachment) ---
            try
            {
                if (mi != null && mi.Attachments != null && mi.Attachments.Count > 0)
                {
                    string __tmpDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "CrmRegardingAddin_Attach_" + System.Guid.NewGuid().ToString("N"));
                    System.IO.Directory.CreateDirectory(__tmpDir);
                    try
                    {
                        foreach (Outlook.Attachment att in mi.Attachments)
                        {
                            string fn = null;
                            try { fn = att.FileName; } catch { }
                            if (string.IsNullOrWhiteSpace(fn)) { try { fn = att.DisplayName; } catch { }
                            if (string.IsNullOrWhiteSpace(fn)) fn = "attachment.bin"; }
                            string savePath = System.IO.Path.Combine(__tmpDir, fn);
                            try { att.SaveAsFile(savePath); } catch { continue; }
                            byte[] bytes = null;
                            try { bytes = System.IO.File.ReadAllBytes(savePath); } catch { bytes = null; }
                            if (bytes == null || bytes.Length == 0) { try { System.IO.File.Delete(savePath); } catch { } continue; }

                            string mime = InferMimeTypeFromFileName(fn);

                            var a = new Microsoft.Xrm.Sdk.Entity("activitymimeattachment");
                            a["objectid"] = new Microsoft.Xrm.Sdk.EntityReference("email", id);
                            a["objecttypecode"] = "email";
                            a["subject"] = fn;
                            a["filename"] = fn;
                            if (!string.IsNullOrWhiteSpace(mime)) a["mimetype"] = mime;
                            a["body"] = System.Convert.ToBase64String(bytes);

                            try { org.Create(a); } catch { }

                            try { System.IO.File.Delete(savePath); } catch { }
                        }
                    }
                    finally
                    {
                        try { System.IO.Directory.Delete(__tmpDir, true); } catch { }
                    }
                }
            }
            catch { }
            TrySetActivityCompleted(org, "email", id);
            return id;
        }

        // ======================== FINALIZE (Outlook store) ========================

        public static void FinalizeAppointmentLinkInOutlookStoreAfterCrmCommit(IOrganizationService org, Outlook.AppointmentItem appt, Guid createdId)
        {
            if (appt == null) return;
            var pa = appt.PropertyAccessor;

            var regId  = TryGet(pa, DaslId.RegardingId, DaslStr.RegardingId_Lower, DaslStr.RegardingId_Camel, DaslStr.RegardingId_Old);
            var regOTs = TryGet(pa, DaslId.RegardingObjectType, DaslStr.RegardingTypeCode_Lower, DaslStr.RegardingTypeCode_Camel, DaslStr.RegardingType_Old);
            Guid regardingId; Guid.TryParse((regId ?? "").Trim('{','}'), out regardingId);
            int typeCode;
            if (!int.TryParse(regOTs ?? "", NumberStyles.Integer, CultureInfo.InvariantCulture, out typeCode))
                typeCode = (ResolveTypeCode(org, regOTs) ?? 4201);

            var orgId = GetOrgId(org);
            string entryId = null; try { entryId = appt.EntryID; } catch { }

            CrmMapiInterop.SetCrmLinkPropsForAppointment(appt, regardingId, typeCode, createdId, orgId, entryId);

            // Mirror "OK" CSV props: 0x80F3, 0x80F2, crmObjectTypeCode, owner, linkstate=2 (DOUBLE)
            try
            {
                string br = Braced(regardingId);
                if (!string.IsNullOrWhiteSpace(br))
                {
                    TrySet(pa, DaslStr.RegardingId_Old, br);
                    TrySet(pa, DaslStr.RegardingId_Lower, br);
                    TrySet(pa, DaslStr.RegardingId_Camel, br);
                }
                if (typeCode > 0)
                {
                    string otcStr = typeCode.ToString(CultureInfo.InvariantCulture);
                    TrySet(pa, DaslStr.RegardingType_Old, otcStr);           // 0x80F2
                    TrySet(pa, DaslStr.RegardingTypeCode_Lower, otcStr);
                    TrySet(pa, DaslStr.RegardingTypeCode_Camel, otcStr);
                }
                if (orgId.HasValue) { TrySet(pa, DaslId.OrgId, Braced(orgId.Value)); TrySet(pa, DaslStr.OrgId, Braced(orgId.Value)); }

                // owner
                try
                {
                    var who = (WhoAmIResponse)org.Execute(new WhoAmIRequest());
                    if (who != null && who.UserId != Guid.Empty)
                    {
                        TrySet(pa, DaslStr.OwnerId, Braced(who.UserId));
                        TrySet(pa, DaslStr.OwnerIdType, OTC_SYSTEMUSER.ToString(CultureInfo.InvariantCulture));
                    }
                } catch { }

                TrySet(pa, DaslId.ObjectTypeCode, (double)4201);
                TrySet(pa, DaslStr.ObjectTypeCode, "4201");

                TrySet(pa, DaslId.LinkState, 2.0);
                TrySet(pa, DaslStr.LinkState, "2");
            }
            catch { }

            try { appt.Save(); } catch { }
        }

        public static void FinalizeMailLinkInOutlookStoreAfterCrmCommit(IOrganizationService org, Outlook.MailItem mi, Guid createdEmailId)
        {
            if (mi == null) return; var pa = mi.PropertyAccessor;
            var orgId = GetOrgId(org);
            var regardingIdStr = TryGet(pa, DaslId.RegardingId, DaslStr.RegardingId_Camel, DaslStr.RegardingId_Lower, DaslStr.RegardingId_Old);
            var regardingLogical = TryGet(pa, DaslId.RegardingObjectType, DaslStr.RegardingType_Old);
            TrySet(pa, DaslId.CrmId, Braced(createdEmailId));        TrySet(pa, DaslStr.CrmId, Braced(createdEmailId));
            if (orgId.HasValue) { TrySet(pa, DaslId.OrgId, Braced(orgId.Value)); TrySet(pa, DaslStr.OrgId, Braced(orgId.Value)); }
            TrySet(pa, DaslId.LinkState, 2.0);                       TrySet(pa, DaslStr.LinkState, "2");
            if (!string.IsNullOrWhiteSpace(regardingIdStr)) { TrySet(pa, DaslId.RegardingId, regardingIdStr); TrySet(pa, DaslStr.RegardingId_Camel, regardingIdStr); TrySet(pa, DaslStr.RegardingId_Lower, regardingIdStr); TrySet(pa, DaslStr.RegardingId_Old, regardingIdStr); }
            if (!string.IsNullOrWhiteSpace(regardingLogical)) { TrySet(pa, DaslId.RegardingObjectType, regardingLogical); }
            TrySet(pa, DaslId.ObjectTypeCode, (double)4202);         TrySet(pa, DaslStr.ObjectTypeCode, "4202");

            // Build and set CrmPartyInfo for mail
            try
            {
                var xml = BuildPartyInfo(org, mi);
                if (!string.IsNullOrWhiteSpace(xml))
                {
                    TrySet(pa, DaslId.PartyInfo, xml);
                    TrySet(pa, DaslStr.PartyInfo, xml);
                }
            } catch { }

            TrySet(pa, DaslId.AsyncSend, 0.0);                       TrySet(pa, DaslStr.AsyncSend, "0");
            try { var msgId = MailUtil.GetInternetMessageId(mi); if (!string.IsNullOrWhiteSpace(msgId)) { TrySet(pa, DaslId.CrmMessageId, msgId); TrySet(pa, DaslStr.CrmMessageId, msgId); } } catch { }
            TrySet(pa, DaslId.SssPromoteTracker, 1);                 TrySet(pa, DaslStr.SssPromoteTracker, "1");
            TrySet(pa, DaslId.TrackedBySender, false);               TrySet(pa, DaslStr.TrackedBySender, "False");
            try { mi.Save(); } catch { }
        }

        
        private static string UnbracedLower(Guid g) { return g.ToString("D").ToLowerInvariant(); }

        private static string GetCrmDisplayName(IOrganizationService org, string logicalName, Guid id)
        {
            try
            {
                string attr = null;
                switch ((logicalName ?? "").ToLowerInvariant())
                {
                    case "systemuser": attr = "fullname"; break;
                    case "contact":    attr = "fullname"; break;
                    case "lead":       attr = "fullname"; break;
                    case "account":    attr = "name";     break;
                }
                if (string.IsNullOrWhiteSpace(attr)) return null;
                var e = org.Retrieve(logicalName, id, new ColumnSet(attr));
                if (e == null) return null;
                var o = e.Contains(attr) ? e[attr] : null;
                return o != null ? Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture) : null;
            }
            catch { return null; }
        }

        private static Tuple<Guid,int,string> ResolveCrmPartyByEmailWithName(IOrganizationService org, string smtp)
        {
            if (org == null || string.IsNullOrWhiteSpace(smtp)) return null;

            // Priority: systemuser -> contact -> account -> lead
            var e = TryFindEntityByEmail(org, "systemuser", "internalemailaddress", smtp, new []{"fullname"});
            if (e != null) return Tuple.Create(e.Id, OTC_SYSTEMUSER, SafeGetString(e, "fullname"));

            e = TryFindEntityByEmail(org, "contact", "emailaddress1", smtp, new []{"fullname"});
            if (e != null) return Tuple.Create(e.Id, OTC_CONTACT, SafeGetString(e, "fullname"));

            e = TryFindEntityByEmail(org, "account", "emailaddress1", smtp, new []{"name"});
            if (e != null) return Tuple.Create(e.Id, OTC_ACCOUNT, SafeGetString(e, "name"));

            e = TryFindEntityByEmail(org, "lead", "emailaddress1", smtp, new []{"fullname"});
            if (e != null) return Tuple.Create(e.Id, OTC_LEAD, SafeGetString(e, "fullname"));

            return null;
        }

        private static Microsoft.Xrm.Sdk.Entity TryFindEntityByEmail(IOrganizationService org, string logicalName, string attribute, string email, string[] columns)
        {
            try
            {
                var q = new QueryExpression(logicalName) { ColumnSet = new ColumnSet(columns ?? new string[0]), TopCount = 1 };
                q.Criteria.AddCondition(attribute, ConditionOperator.Equal, email);
                var r = org.RetrieveMultiple(q);
                return r.Entities.FirstOrDefault();
            }
            catch { return null; }
        }

        private static string SafeGetString(Microsoft.Xrm.Sdk.Entity e, string attr)
        {
            try
            {
                if (e == null || string.IsNullOrWhiteSpace(attr)) return null;
                if (!e.Contains(attr)) return null;
                var o = e[attr];
                return o != null ? Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture) : null;
            }
            catch { return null; }
        }

        
        private static string GetOutgoingSenderSmtp(Outlook.MailItem mi)
        {
            try
            {
                var acc = mi.SendUsingAccount;
                if (acc != null && !string.IsNullOrWhiteSpace(acc.SmtpAddress))
                    return acc.SmtpAddress;
            } catch { }
            try
            {
                var ns = mi.Session;
                var ae = ns != null ? ns.CurrentUser?.AddressEntry : null;
                if (ae != null)
                {
                    try { var exu = ae.GetExchangeUser(); if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress)) return exu.PrimarySmtpAddress; } catch { }
                    try { var addr = ae.Address; if (!string.IsNullOrWhiteSpace(addr) && addr.Contains("@")) return addr; } catch { }
                }
            } catch { }
            try
            {
                var val = mi.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0065001F") as string;
                if (!string.IsNullOrWhiteSpace(val) && val.Contains("@"))
                    return val;
            } catch { }
            try
            {
                if (!string.IsNullOrWhiteSpace(mi.SenderEmailAddress) && mi.SenderEmailAddress.Contains("@"))
                    return mi.SenderEmailAddress;
            } catch { }
            return null;
        }

        // ======================== Helpers ========================
                    

        private static T Safe<T>(Func<T> f) { try { return f(); } catch { return default(T); } }

        private static Guid FindCrmAppointmentIdByGlobalId(IOrganizationService org, string globalId)
        {
            try
            {
                var q = new QueryExpression("appointment") { ColumnSet = new ColumnSet("activityid") };
                q.Criteria.AddCondition("globalobjectid", ConditionOperator.Equal, globalId);
                var r = org.RetrieveMultiple(q);
                var e = r.Entities.FirstOrDefault();
                return e != null ? e.Id : Guid.Empty;
            } catch { return Guid.Empty; }
        }

        private static void TryBackfillApptPartiesIfEmpty(IOrganizationService org, Guid apptId, Entity update, Outlook.AppointmentItem appt)
        {
            try
            {
                var full = org.Retrieve("appointment", apptId, new ColumnSet("requiredattendees", "optionalattendees"));
                bool hasReq = full.Contains("requiredattendees") && full.GetAttributeValue<EntityCollection>("requiredattendees")?.Entities?.Count > 0;
                bool hasOpt = full.Contains("optionalattendees") && full.GetAttributeValue<EntityCollection>("optionalattendees")?.Entities?.Count > 0;
                if (!hasReq) update["requiredattendees"] = ActivityPartyBuilder.BuildApptRecipientsFromRecipients(org, appt, true);
                if (!hasOpt) update["optionalattendees"] = ActivityPartyBuilder.BuildApptRecipientsFromRecipients(org, appt, false);
            } catch { }
        }

        // Build crmpartyinfo XML for Appointment with PartyId + TypeCode.
        private static string BuildPartyInfo(IOrganizationService org, Outlook.AppointmentItem appt)
        {
            try
            {
                if (appt == null) return null;
                var sb = new StringBuilder();
                sb.Append("<PartyMembers Version=\"1.0\">");

                // Required attendees (5), then Optional (6) — before Organizer (7) to match expected order.
                AppendApptRecipientsWithCrm(org, sb, appt, Outlook.OlMeetingRecipientType.olRequired, "5");
                AppendApptRecipientsWithCrm(org, sb, appt, Outlook.OlMeetingRecipientType.olOptional, "6");

                // Organizer (7) – resolve PartyId (systemuser) from WhoAmI and email from CRM user or Outlook
                try
                {
                    Guid? userId = null; string orgEmail = null; string orgName = null;
                    try
                    {
                        var who = (WhoAmIResponse)org.Execute(new WhoAmIRequest());
                        if (who != null) { userId = who.UserId; }
                        var user = (who != null) ? org.Retrieve("systemuser", who.UserId, new ColumnSet("internalemailaddress", "fullname")) : null;
                        if (user != null)
                        {
                            orgEmail = user.GetAttributeValue<string>("internalemailaddress");
                            orgName  = user.GetAttributeValue<string>("fullname");
                        }
                    } catch { }
                    try { if (string.IsNullOrWhiteSpace(orgName)) orgName = appt.Organizer; } catch { }
                    if (string.IsNullOrWhiteSpace(orgName)) orgName = orgEmail;
                    if (!string.IsNullOrWhiteSpace(orgEmail))
                    {
                        sb.Append("<Member Email=\"").Append(SecurityElement.Escape(orgEmail)).Append("\" ");
                        if (userId.HasValue) sb.Append("PartyId=\"").Append(SecurityElement.Escape(Braced(userId.Value))).Append("\" ");
                        sb.Append("TypeCode=\"").Append(OTC_SYSTEMUSER.ToString(CultureInfo.InvariantCulture)).Append("\" ");
                        sb.Append("Name=\"").Append(SecurityElement.Escape(orgName ?? string.Empty)).Append("\" ParticipationType=\"7\" />");
                    }
                }
                catch { }

                sb.Append("</PartyMembers>");
                return sb.ToString();
            }
            catch { return null; }
        }

        private static void AppendApptRecipientsWithCrm(IOrganizationService org, StringBuilder sb, Outlook.AppointmentItem appt, Outlook.OlMeetingRecipientType rtype, string participation)
        {
            try
            {
                foreach (Outlook.Recipient r in appt.Recipients)
                {
                    if (r == null) continue;
                    if (r.Type != (int)rtype) continue;

                    string smtp = null;
                    string name = null;
                    try { name = r.Name; } catch { }

                    try
                    {
                        if (r.AddressEntry != null)
                        {
                            try
                            {
                                var exu = r.AddressEntry.GetExchangeUser();
                                if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress))
                                {
                                    smtp = exu.PrimarySmtpAddress;
                                    if (string.IsNullOrWhiteSpace(name)) name = exu.Name;
                                }
                                else
                                {
                                    var dl = r.AddressEntry.GetExchangeDistributionList();
                                    if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress))
                                    {
                                        smtp = dl.PrimarySmtpAddress;
                                        if (string.IsNullOrWhiteSpace(name)) name = dl.Name;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }

                    if (string.IsNullOrWhiteSpace(smtp)) { try { smtp = r.Address; } catch { } }
                    if (string.IsNullOrWhiteSpace(name))  { name = smtp; }
                    if (string.IsNullOrWhiteSpace(smtp)) continue;

                    Guid? pid = null; int? otc = null;
                    try
                    {
                        var resolved = ResolveCrmPartyByEmail(org, smtp);
                        if (resolved != null) { pid = resolved.Item1; otc = resolved.Item2; }
                    }
                    catch { }

                    sb.Append("<Member Email=\"").Append(SecurityElement.Escape(smtp)).Append("\" ");
                    if (pid.HasValue) sb.Append("PartyId=\"").Append(SecurityElement.Escape(Braced(pid.Value))).Append("\" ");
                    if (otc.HasValue) sb.Append("TypeCode=\"").Append(otc.Value.ToString(CultureInfo.InvariantCulture)).Append("\" ");
                    sb.Append("Name=\"").Append(SecurityElement.Escape(name ?? string.Empty)).Append("\" ParticipationType=\"")
                      .Append(participation).Append("\" />");
                }
            }
            catch { }
        }

        // NEW: Build crmpartyinfo XML for Mail with PartyId + TypeCode.
        
        // Build crmpartyinfo XML for Mail (MS add-in compatible):
        // - No ParticipationType attribute
        // - Order: From first, then To, Cc, Bcc
        // - Deduplicate by email (case-insensitive)
        // - PartyId = guid sans accolades en minuscules ; TypeCode = OTC (8/2/1/4). If unresolved: PartyId="", TypeCode="-1"
        // - Name = CRM display name if resolved; otherwise email
        private static string BuildPartyInfo(IOrganizationService org, Outlook.MailItem mi)
        {
            try
            {
                if (mi == null) return null;
                var sb = new StringBuilder();
                sb.Append("<PartyMembers Version=\"1.0\">");

                var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

                // Helper local function to append one member
                Action<string,string,Guid?,int?,string> append = (smtp, fallbackName, pid, otc, resolvedName) =>
                {
                    if (string.IsNullOrWhiteSpace(smtp)) return;
                    if (seen.Contains(smtp)) return;
                    seen.Add(smtp);

                    string name = !string.IsNullOrWhiteSpace(resolvedName) ? resolvedName : (fallbackName ?? smtp);
                    string partyIdStr = pid.HasValue ? UnbracedLower(pid.Value) : "";
                    string typeCodeStr = otc.HasValue ? otc.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "-1";

                    sb.Append("<Member Email=\"").Append(SecurityElement.Escape(smtp)).Append("\" ");
                    sb.Append("PartyId=\"").Append(SecurityElement.Escape(partyIdStr)).Append("\" ");
                    sb.Append("TypeCode=\"").Append(typeCodeStr).Append("\" ");
                    sb.Append("Name=\"").Append(SecurityElement.Escape(name ?? string.Empty)).Append("\" />");
                };

                // 1) From (first)
                try
                {
                    string fromSmtp = GetOutgoingSenderSmtp(mi);
                    if (string.IsNullOrWhiteSpace(fromSmtp)) fromSmtp = ActivityPartyBuilder.GetSenderSmtp(mi);
                    string fromName = null;
                    try { var acc = mi.SendUsingAccount; if (acc!=null && !string.IsNullOrWhiteSpace(acc.DisplayName)) fromName = acc.DisplayName; } catch { }
                    if (string.IsNullOrWhiteSpace(fromName)) { try { fromName = mi.SenderName; } catch { } }

                    Guid? pid = null; int? otc = null; string display = null;
                    try
                    {
                        var resolved = ResolveCrmPartyByEmailWithName(org, fromSmtp);
                        if (resolved != null) { pid = resolved.Item1; otc = resolved.Item2; display = resolved.Item3; }
                    } catch { }
                    append(fromSmtp, fromName, pid, otc, display);
                }
                catch { }

                // 2) To, then Cc, then Bcc
                try
                {
                    var order = new [] {
                        Outlook.OlMailRecipientType.olTo,
                        Outlook.OlMailRecipientType.olCC,
                        Outlook.OlMailRecipientType.olBCC
                    };

                    foreach (var type in order)
                    {
                        foreach (Outlook.Recipient r in mi.Recipients)
                        {
                            if (r == null || r.Type != (int)type) continue;

                            string smtp = null;
                            string name = null;
                            try { name = r.Name; } catch { }

                            try
                            {
                                if (r.AddressEntry != null)
                                {
                                    var exu = r.AddressEntry.GetExchangeUser();
                                    if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress))
                                    {
                                        smtp = exu.PrimarySmtpAddress;
                                        if (string.IsNullOrWhiteSpace(name)) name = exu.Name;
                                    }
                                    else
                                    {
                                        var dl = r.AddressEntry.GetExchangeDistributionList();
                                        if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress))
                                        {
                                            smtp = dl.PrimarySmtpAddress;
                                            if (string.IsNullOrWhiteSpace(name)) name = dl.Name;
                                        }
                                    }
                                }
                            }
                            catch { }
                            if (string.IsNullOrWhiteSpace(smtp)) { try { smtp = r.Address; } catch { } }
                            if (string.IsNullOrWhiteSpace(smtp)) continue;
                            if (string.IsNullOrWhiteSpace(name))  { name = smtp; }

                            Guid? pid = null; int? otc = null; string display = null;
                            try
                            {
                                var resolved = ResolveCrmPartyByEmailWithName(org, smtp);
                                if (resolved != null) { pid = resolved.Item1; otc = resolved.Item2; display = resolved.Item3; }
                            }
                            catch { }

                            append(smtp, name, pid, otc, display);
                        }
                    }
                }
                catch { }

                sb.Append("</PartyMembers>");
                return sb.ToString();
            }
            catch { return null; }
        }


        private static void AppendMailFromWithCrm(IOrganizationService org, StringBuilder sb, Outlook.MailItem mi, string participation)
        {
            try
            {
                string smtp = ActivityPartyBuilder.GetSenderSmtp(mi);
                string name = null;
                try { name = mi.SenderName; } catch { }
                try
                {
                    var ae = mi.Sender;
                    if (ae != null)
                    {
                        var exu = ae.GetExchangeUser();
                        if (exu != null)
                        {
                            if (string.IsNullOrWhiteSpace(smtp)) smtp = exu.PrimarySmtpAddress;
                            if (string.IsNullOrWhiteSpace(name)) name = exu.Name;
                        }
                        else
                        {
                            var dl = ae.GetExchangeDistributionList();
                            if (dl != null)
                            {
                                if (string.IsNullOrWhiteSpace(smtp)) smtp = dl.PrimarySmtpAddress;
                                if (string.IsNullOrWhiteSpace(name)) name = dl.Name;
                            }
                        }
                    }
                }
                catch { }
                if (string.IsNullOrWhiteSpace(name)) name = smtp;
                if (string.IsNullOrWhiteSpace(smtp)) return;

                Guid? pid = null; int? otc = null;
                try
                {
                    var resolved = ResolveCrmPartyByEmail(org, smtp);
                    if (resolved != null) { pid = resolved.Item1; otc = resolved.Item2; }
                }
                catch { }

                sb.Append("<Member Email=\"").Append(SecurityElement.Escape(smtp)).Append("\" ");
                if (pid.HasValue) sb.Append("PartyId=\"").Append(SecurityElement.Escape(Braced(pid.Value))).Append("\" ");
                if (otc.HasValue) sb.Append("TypeCode=\"").Append(otc.Value.ToString(CultureInfo.InvariantCulture)).Append("\" ");
                sb.Append("Name=\"").Append(SecurityElement.Escape(name ?? string.Empty)).Append("\" ParticipationType=\"")
                  .Append(participation).Append("\" />");
            }
            catch { }
        }

        private static void AppendMailRecipientsWithCrm(IOrganizationService org, StringBuilder sb, Outlook.MailItem mi, Outlook.OlMailRecipientType type, string participation)
        {
            try
            {
                foreach (Outlook.Recipient r in mi.Recipients)
                {
                    if (r.Type != (int)type) continue;

                    string smtp = null;
                    string name = null;
                    try { name = r.Name; } catch { }

                    try
                    {
                        if (r.AddressEntry != null)
                        {
                            var exu = r.AddressEntry.GetExchangeUser();
                            if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress))
                            {
                                smtp = exu.PrimarySmtpAddress;
                                if (string.IsNullOrWhiteSpace(name)) name = exu.Name;
                            }
                            else
                            {
                                var dl = r.AddressEntry.GetExchangeDistributionList();
                                if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress))
                                {
                                    smtp = dl.PrimarySmtpAddress;
                                    if (string.IsNullOrWhiteSpace(name)) name = dl.Name;
                                }
                            }
                        }
                    }
                    catch { }
                    if (string.IsNullOrWhiteSpace(smtp)) { try { smtp = r.Address; } catch { } }
                    if (string.IsNullOrWhiteSpace(name))  { name = smtp; }
                    if (string.IsNullOrWhiteSpace(smtp)) continue;

                    Guid? pid = null; int? otc = null;
                    try
                    {
                        var resolved = ResolveCrmPartyByEmail(org, smtp);
                        if (resolved != null) { pid = resolved.Item1; otc = resolved.Item2; }
                    }
                    catch { }

                    sb.Append("<Member Email=\"").Append(SecurityElement.Escape(smtp)).Append("\" ");
                    if (pid.HasValue) sb.Append("PartyId=\"").Append(SecurityElement.Escape(Braced(pid.Value))).Append("\" ");
                    if (otc.HasValue) sb.Append("TypeCode=\"").Append(otc.Value.ToString(CultureInfo.InvariantCulture)).Append("\" ");
                    sb.Append("Name=\"").Append(SecurityElement.Escape(name ?? string.Empty)).Append("\" ParticipationType=\"")
                      .Append(participation).Append("\" />");
                }
            }
            catch { }
        }

        private static Tuple<Guid,int> ResolveCrmPartyByEmail(IOrganizationService org, string smtp)
        {
            if (org == null || string.IsNullOrWhiteSpace(smtp)) return null;
            // priority: systemuser (internalemailaddress), contact, account, lead
            var found = TryFindByEmail(org, "systemuser", "internalemailaddress", smtp);
            if (found != null) return Tuple.Create(found.Id, OTC_SYSTEMUSER);
            found = TryFindByEmail(org, "contact", "emailaddress1", smtp);
            if (found != null) return Tuple.Create(found.Id, ResolveTypeCode(org, "contact") ?? 2);
            found = TryFindByEmail(org, "account", "emailaddress1", smtp);
            if (found != null) return Tuple.Create(found.Id, ResolveTypeCode(org, "account") ?? 1);
            found = TryFindByEmail(org, "lead", "emailaddress1", smtp);
            if (found != null) return Tuple.Create(found.Id, ResolveTypeCode(org, "lead") ?? 4);
            return null;
        }

        private static EntityReference TryFindByEmail(IOrganizationService org, string logicalName, string attribute, string email)
        {
            try
            {
                var q = new QueryExpression(logicalName) { ColumnSet = new ColumnSet(false), TopCount = 1 };
                q.Criteria.AddCondition(attribute, ConditionOperator.Equal, email);
                var r = org.RetrieveMultiple(q);
                var e = r.Entities.FirstOrDefault();
                return e != null ? new EntityReference(logicalName, e.Id) : null;
            }
            catch { return null; }
        }

        private static void TrySetActivityCompleted(IOrganizationService org, string logicalName, Guid id)
        {
            try
            {
                var attrReq = new RetrieveAttributeRequest
                {
                    EntityLogicalName = logicalName,
                    LogicalName = "statuscode",
                    RetrieveAsIfPublished = true
                };
                var attrResp = (RetrieveAttributeResponse)org.Execute(attrReq);
                var statusMeta = attrResp.AttributeMetadata as StatusAttributeMetadata;
                int statusForCompleted = 0;
                if (statusMeta != null)
                {
                    var opt = statusMeta.OptionSet.Options
                        .OfType<StatusOptionMetadata>();
                    var firstCompleted = opt.FirstOrDefault(o => o.State == 1);
                    if (firstCompleted != null && firstCompleted.Value.HasValue) statusForCompleted = firstCompleted.Value.Value;
                }
                if (statusForCompleted == 0) statusForCompleted = 2;

                var set = new SetStateRequest
                {
                    EntityMoniker = new EntityReference(logicalName, id),
                    State = new OptionSetValue(1),
                    Status = new OptionSetValue(statusForCompleted)
                };
                org.Execute(set);
            }
            catch { }
        }

        private static class ActivityPartyBuilder
        {
            public static string GetSenderSmtp(Outlook.MailItem mi)
            {
                try
                {
                    if (mi == null) return null;
                    try { var smtpPref = GetOutgoingSenderSmtp(mi); if (!string.IsNullOrWhiteSpace(smtpPref)) return smtpPref; } catch { }

                    Outlook.AddressEntry s = mi.Sender;
                    if (s != null)
                    {
                        try { var exu = s.GetExchangeUser(); if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress)) return exu.PrimarySmtpAddress; } catch { }
                        try { var dl  = s.GetExchangeDistributionList(); if (dl  != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress))  return dl.PrimarySmtpAddress; } catch { }
                        try { var addr = s.Address; if (!string.IsNullOrWhiteSpace(addr)) return addr; } catch { }
                    }
                    var raw = mi.SenderEmailAddress;
                    if (!string.IsNullOrWhiteSpace(raw)) return raw;
                }
                catch { }
                return null;
            }

            public static string GetCurrentUserEmail(IOrganizationService org)
            {
                try
                {
                    var who = (WhoAmIResponse)org.Execute(new WhoAmIRequest());
                    if (who == null) return null;
                    var user = org.Retrieve("systemuser", who.UserId, new ColumnSet("internalemailaddress"));
                    var mail = user.GetAttributeValue<string>("internalemailaddress");
                    return string.IsNullOrWhiteSpace(mail) ? null : mail;
                }
                catch { return null; }
            }

            public static EntityCollection BuildFromForMail(IOrganizationService org, Outlook.MailItem mi)
            {
                var parts = new EntityCollection { EntityName = "activityparty" };
                try
                {
                    var smtp = GetSenderSmtp(mi);
                    if (string.IsNullOrWhiteSpace(smtp)) return parts;

                    var er = TryFindByEmail(org, "systemuser", "internalemailaddress", smtp)
                          ?? TryFindByEmail(org, "contact", "emailaddress1", smtp)
                          ?? TryFindByEmail(org, "account", "emailaddress1", smtp)
                          ?? TryFindByEmail(org, "lead", "emailaddress1", smtp);

                    var p = new Entity("activityparty");
                    if (er != null) p["partyid"] = er; else p["addressused"] = smtp;
                    p["participationtypemask"] = new OptionSetValue(1); // From
                    parts.Entities.Add(p);
                }
                catch { }
                return parts;
            }

            public static EntityCollection BuildRecipientsFromOutlook(IOrganizationService org, Outlook.MailItem mi, Outlook.OlMailRecipientType type)
            {
                var parts = new EntityCollection { EntityName = "activityparty" };
                try
                {
                    foreach (Outlook.Recipient r in mi.Recipients)
                    {
                        if (r.Type != (int)type) continue;
                        string smtp = null;
                        try
                        {
                            if (r.AddressEntry != null)
                            {
                                var exu = r.AddressEntry.GetExchangeUser();
                                if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress)) smtp = exu.PrimarySmtpAddress;
                                else
                                {
                                    var dl = r.AddressEntry.GetExchangeDistributionList();
                                    if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress)) smtp = dl.PrimarySmtpAddress;
                                }
                            }
                        }
                        catch { }
                        if (string.IsNullOrWhiteSpace(smtp)) { try { smtp = r.Address; } catch { } }
                        if (string.IsNullOrWhiteSpace(smtp)) continue;
                        var er = TryFindByEmail(org, "systemuser", "internalemailaddress", smtp)
                             ?? TryFindByEmail(org, "contact", "emailaddress1", smtp)
                             ?? TryFindByEmail(org, "account", "emailaddress1", smtp)
                             ?? TryFindByEmail(org, "lead", "emailaddress1", smtp);
                        var p = new Entity("activityparty"); if (er != null) p["partyid"] = er; else p["addressused"] = smtp;
                        int mask = type == Outlook.OlMailRecipientType.olTo ? 2 :
                                   type == Outlook.OlMailRecipientType.olCC ? 3 : 4;
                        p["participationtypemask"] = new OptionSetValue(mask);
                        parts.Entities.Add(p);
                    }
                }
                catch { }
                return parts;
            }

            public static EntityCollection BuildApptRecipientsFromRecipients(IOrganizationService org, Outlook.AppointmentItem appt, bool required)
            {
                var parts = new EntityCollection { EntityName = "activityparty" };
                try
                {
                    foreach (Outlook.Recipient r in appt.Recipients)
                    {
                        int rtype = r.Type; // OlMeetingRecipientType
                        bool isReq = (rtype == (int)Outlook.OlMeetingRecipientType.olRequired);
                        bool isOpt = (rtype == (int)Outlook.OlMeetingRecipientType.olOptional);
                        if (required && !isReq) continue;
                        if (!required && !isOpt) continue;

                        string smtp = null;
                        try
                        {
                            if (r.AddressEntry != null)
                            {
                                var exu = r.AddressEntry.GetExchangeUser();
                                if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress)) smtp = exu.PrimarySmtpAddress;
                                else
                                {
                                    var dl = r.AddressEntry.GetExchangeDistributionList();
                                    if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress)) smtp = dl.PrimarySmtpAddress;
                                }
                            }
                        }
                        catch { }
                        if (string.IsNullOrWhiteSpace(smtp)) { try { smtp = r.Address; } catch { } }
                        if (string.IsNullOrWhiteSpace(smtp)) continue;

                        var er = TryFindByEmail(org, "systemuser", "internalemailaddress", smtp)
                             ?? TryFindByEmail(org, "contact", "emailaddress1", smtp)
                             ?? TryFindByEmail(org, "account", "emailaddress1", smtp)
                             ?? TryFindByEmail(org, "lead", "emailaddress1", smtp);

                        var p = new Entity("activityparty");
                        if (er != null) p["partyid"] = er; else p["addressused"] = smtp;
                        p["participationtypemask"] = new OptionSetValue(required ? 5 : 6);
                        parts.Entities.Add(p);
                    }
                }
                catch { }
                return parts;
            }

            public static EntityCollection BuildOrganizer(IOrganizationService org)
            {
                var parts = new EntityCollection { EntityName = "activityparty" };
                try
                {
                    var who = (WhoAmIResponse)org.Execute(new WhoAmIRequest());
                    if (who != null)
                    {
                        var p = new Entity("activityparty");
                        p["partyid"] = new EntityReference("systemuser", who.UserId);
                        p["participationtypemask"] = new OptionSetValue(7); // Organizer
                        parts.Entities.Add(p);
                    }
                }
                catch { }
                return parts;
            }

            private static EntityReference TryFindByEmail(IOrganizationService org, string logicalName, string attribute, string email)
            {
                try
                {
                    var q = new QueryExpression(logicalName) { ColumnSet = new ColumnSet(false), TopCount = 1 };
                    q.Criteria.AddCondition(attribute, ConditionOperator.Equal, email);
                    var r = org.RetrieveMultiple(q);
                    var e = r.Entities.FirstOrDefault();
                    return e != null ? new EntityReference(logicalName, e.Id) : null;
                }
                catch { return null; }
            }
        }

        // Helper MIME: used for uploading Outlook attachments to CRM
        private static string InferMimeTypeFromFileName(string fileName)
        {
            try
            {
                var ext = System.IO.Path.GetExtension(fileName);
                if (string.IsNullOrWhiteSpace(ext))
                    return "application/octet-stream";

                ext = ext.Trim().Trim('.').ToLowerInvariant();
                switch (ext)
                {
                    case "txt": return "text/plain";
                    case "html":
                    case "htm": return "text/html";
                    case "csv": return "text/csv";
                    case "xml": return "application/xml";
                    case "json": return "application/json";
                    case "pdf": return "application/pdf";

                    case "doc": return "application/msword";
                    case "docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    case "xls": return "application/vnd.ms-excel";
                    case "xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    case "ppt": return "application/vnd.ms-powerpoint";
                    case "pptx": return "application/vnd.openxmlformats-officedocument.presentationml.presentation";

                    case "jpg":
                    case "jpeg": return "image/jpeg";
                    case "png": return "image/png";
                    case "gif": return "image/gif";
                    case "bmp": return "image/bmp";
                    case "tif":
                    case "tiff": return "image/tiff";

                    case "zip": return "application/zip";
                    case "rar": return "application/x-rar-compressed";
                    case "7z": return "application/x-7z-compressed";

                    case "eml": return "message/rfc822";
                }
            }
            catch { /* ignore and fall back */ }

            return "application/octet-stream";
        }

        // === Ensure Outlook's UserProperties (0x80C8) is present and defines MS CRM fields ===
        private static void EnsureUserPropertiesForMail(Outlook.MailItem mi)
        {
            if (mi == null) return;
            try
            {
                var pa = mi.PropertyAccessor;
                var ups = mi.UserProperties;
                EnsureUserProperty(ups, "crmLinkState", Outlook.OlUserPropertyType.olNumber, TryGet(pa, DaslId.LinkState, DaslStr.LinkState));
                EnsureUserProperty(ups, "crmRegardingObjectId", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslId.RegardingId, DaslStr.RegardingId_Camel, DaslStr.RegardingId_Lower, DaslStr.RegardingId_Old));
                EnsureUserProperty(ups, "crmRegardingObjectTypeCode", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslId.RegardingObjectType, DaslStr.RegardingTypeCode_Camel, DaslStr.RegardingTypeCode_Lower, DaslStr.RegardingType_Old));
                EnsureUserProperty(ups, "Regarding", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslStr.RegardingLabel));
                EnsureUserProperty(ups, "crmpartyinfo", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslId.PartyInfo, DaslStr.PartyInfo));
                // Do not add to folder fields; Save to force 0x80C8 to be (re)generated by Outlook.
                try { mi.Save(); } catch { }
            }
            catch { }
        }

        private static void EnsureUserPropertiesForAppointment(Outlook.AppointmentItem appt)
        {
            if (appt == null) return;
            try
            {
                var pa = appt.PropertyAccessor;
                var ups = appt.UserProperties;
                EnsureUserProperty(ups, "crmLinkState", Outlook.OlUserPropertyType.olNumber, TryGet(pa, DaslId.LinkState, DaslStr.LinkState));
                EnsureUserProperty(ups, "crmObjectTypeCode", Outlook.OlUserPropertyType.olNumber, TryGet(pa, DaslId.ObjectTypeCode, DaslStr.ObjectTypeCode));
                EnsureUserProperty(ups, "crmRegardingObjectId", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslId.RegardingId, DaslStr.RegardingId_Camel, DaslStr.RegardingId_Lower, DaslStr.RegardingId_Old));
                EnsureUserProperty(ups, "crmRegardingObjectTypeCode", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslId.RegardingObjectType, DaslStr.RegardingTypeCode_Camel, DaslStr.RegardingTypeCode_Lower, DaslStr.RegardingType_Old));
                EnsureUserProperty(ups, "Regarding", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslStr.RegardingLabel));
                EnsureUserProperty(ups, "crmpartyinfo", Outlook.OlUserPropertyType.olText, TryGet(pa, DaslId.PartyInfo, DaslStr.PartyInfo));
                try { appt.Save(); } catch { }
            }
            catch { }
        }

        private static void EnsureUserProperty(Outlook.UserProperties ups, string name, Outlook.OlUserPropertyType type, string value)
        {
            try
            {
                var up = ups.Find(name);
                if (up == null)
                {
                    up = ups.Add(name, type, false /*AddToFolderFields*/);
                }
                if (value != null)
                {
                    try { up.Value = value; } catch { /* type mismatch: ignore */ }
                }
            }
            catch { }
        }

    }
}
