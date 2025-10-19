using System;
using System.Globalization;
using System.Linq;
using System.Text;
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
// === MS add-in string-named aliases for 'regarding' (appointments) ===
private const string DASL_RegId_String_Lower   = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmregardingobjectid";
private const string DASL_RegId_String_Camel   = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmRegardingObjectId";
private const string DASL_RegId_String_Old     = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmRegardingId";
private const string DASL_RegTypeCode_Lower    = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmregardingobjecttypecode";
private const string DASL_RegTypeCode_Camel    = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmRegardingObjectTypeCode";

        private static class DaslId
        {
            public const string LinkState            = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C8"; // PT_DOUBLE
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
            public const string RegardingTypeCode_Lower     = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingobjecttypecode";
            public const string RegardingTypeCode_Camel     = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectTypeCode";
            // Old variants we used previously (fallback read only)
            public const string RegardingId_Old             = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingId";
            public const string RegardingType_Old           = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectType";
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


// === MS add-in regarding aliases after commit (no external deps) ===
try
{
    var __pa2 = appt.PropertyAccessor;
    string __regId2 = null;
    try { __regId2 = System.Convert.ToString(__pa2.GetProperty(DaslId.RegardingId), System.Globalization.CultureInfo.InvariantCulture); } catch {}
    if (string.IsNullOrWhiteSpace(__regId2)) { try { __regId2 = System.Convert.ToString(__pa2.GetProperty(DASL_RegId_String_Lower)); } catch {} }
    if (string.IsNullOrWhiteSpace(__regId2)) { try { __regId2 = System.Convert.ToString(__pa2.GetProperty(DASL_RegId_String_Camel)); } catch {} }
    if (string.IsNullOrWhiteSpace(__regId2)) { try { __regId2 = System.Convert.ToString(__pa2.GetProperty(DASL_RegId_String_Old)); } catch {} }
    if (!string.IsNullOrWhiteSpace(__regId2))
    {
        try { __pa2.SetProperty(DASL_RegId_String_Lower, __regId2); } catch {}
        try { __pa2.SetProperty(DASL_RegId_String_Camel, __regId2); } catch {}
        try { __pa2.SetProperty(DASL_RegId_String_Old,   __regId2); } catch {}
    }
    int __typeCode2 = 0;
    try { var s = System.Convert.ToString(__pa2.GetProperty(DASL_RegTypeCode_Lower), System.Globalization.CultureInfo.InvariantCulture); int.TryParse(s, out __typeCode2); } catch {}
    if (__typeCode2 == 0) { try { var s2 = System.Convert.ToString(__pa2.GetProperty(DASL_RegTypeCode_Camel), System.Globalization.CultureInfo.InvariantCulture); int.TryParse(s2, out __typeCode2); } catch {} }
    if (__typeCode2 == 0) { try { var s3 = System.Convert.ToString(__pa2.GetProperty(DaslId.ObjectTypeCode), System.Globalization.CultureInfo.InvariantCulture); int.TryParse(s3, out __typeCode2); } catch {} }
    if (__typeCode2 > 0)
    {
        var __otcStr2 = __typeCode2.ToString(System.Globalization.CultureInfo.InvariantCulture);
        try { __pa2.SetProperty(DASL_RegTypeCode_Lower, __otcStr2); } catch {}
        try { __pa2.SetProperty(DASL_RegTypeCode_Camel, __otcStr2); } catch {}
    }
}
catch {}


// === MS add-in regarding aliases (no external deps) ===
try
{
    var __pa = appt.PropertyAccessor;
    string __regId = null;
    try { __regId = System.Convert.ToString(__pa.GetProperty(DaslId.RegardingId), System.Globalization.CultureInfo.InvariantCulture); } catch {}
    if (string.IsNullOrWhiteSpace(__regId)) { try { __regId = System.Convert.ToString(__pa.GetProperty(DASL_RegId_String_Lower)); } catch {} }
    if (string.IsNullOrWhiteSpace(__regId)) { try { __regId = System.Convert.ToString(__pa.GetProperty(DASL_RegId_String_Camel)); } catch {} }
    if (string.IsNullOrWhiteSpace(__regId)) { try { __regId = System.Convert.ToString(__pa.GetProperty(DASL_RegId_String_Old)); } catch {} }
    if (!string.IsNullOrWhiteSpace(__regId))
    {
        try { __pa.SetProperty(DASL_RegId_String_Lower, __regId); } catch {}
        try { __pa.SetProperty(DASL_RegId_String_Camel, __regId); } catch {}
        try { __pa.SetProperty(DASL_RegId_String_Old,   __regId); } catch {}
    }
    // type code (if already available)
    int __otc = 0;
    try { var s = System.Convert.ToString(__pa.GetProperty(DASL_RegTypeCode_Lower), System.Globalization.CultureInfo.InvariantCulture); int.TryParse(s, out __otc); } catch {}
    if (__otc == 0) { try { var s2 = System.Convert.ToString(__pa.GetProperty(DASL_RegTypeCode_Camel), System.Globalization.CultureInfo.InvariantCulture); int.TryParse(s2, out __otc); } catch {} }
    if (__otc > 0)
    {
        var __otcStr = __otc.ToString(System.Globalization.CultureInfo.InvariantCulture);
        try { __pa.SetProperty(DASL_RegTypeCode_Lower, __otcStr); } catch {}
        try { __pa.SetProperty(DASL_RegTypeCode_Camel, __otcStr); } catch {}
    }
}
catch {}
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
    }
}
catch { }

            TrySet(pa, DaslId.LinkState, 1.0);                         TrySet(pa, DaslStr.LinkState, "1");
            TrySet(pa, DaslId.RegardingId, Braced(regardingId));       TrySet(pa, DaslStr.RegardingId_Camel, Braced(regardingId)); TrySet(pa, DaslStr.RegardingId_Lower, Braced(regardingId));
            TrySet(pa, DaslId.RegardingObjectType, regardingLogicalName); TrySet(pa, DaslStr.RegardingType_Old, regardingLogicalName);
            if (!string.IsNullOrWhiteSpace(regardingReadableName)) { TrySet(pa, DaslStr.RegardingLabel, regardingReadableName); }
            TrySet(pa, DaslId.SssPromoteTracker, 1);                   TrySet(pa, DaslStr.SssPromoteTracker, "1");
            try { mi.Save(); } catch { }
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
            e["regardingobjectid"] = new EntityReference(regardingLogical, regardingId);

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
            if (!string.IsNullOrWhiteSpace(regardingIdStr)) { TrySet(pa, DaslId.RegardingId, regardingIdStr); TrySet(pa, DaslStr.RegardingId_Camel, regardingIdStr); TrySet(pa, DaslStr.RegardingId_Lower, regardingIdStr); }
            if (!string.IsNullOrWhiteSpace(regardingLogical)) { TrySet(pa, DaslId.RegardingObjectType, regardingLogical); TrySet(pa, DaslStr.RegardingType_Old, regardingLogical); }
            TrySet(pa, DaslId.ObjectTypeCode, (double)4202);         TrySet(pa, DaslStr.ObjectTypeCode, "4202");
            // BuildPartyInfo restored
            var pi = BuildPartyInfo(mi);
            if (!string.IsNullOrWhiteSpace(pi)) { TrySet(pa, DaslId.PartyInfo, pi); TrySet(pa, DaslStr.PartyInfo, pi); }
            TrySet(pa, DaslId.AsyncSend, 0.0);                       TrySet(pa, DaslStr.AsyncSend, "0");
            try { var msgId = MailUtil.GetInternetMessageId(mi); if (!string.IsNullOrWhiteSpace(msgId)) { TrySet(pa, DaslId.CrmMessageId, msgId); TrySet(pa, DaslStr.CrmMessageId, msgId); } } catch { }
            TrySet(pa, DaslId.SssPromoteTracker, 1);                 TrySet(pa, DaslStr.SssPromoteTracker, "1");
            TrySet(pa, DaslId.TrackedBySender, false);               TrySet(pa, DaslStr.TrackedBySender, "False");
            try { mi.Save(); } catch { }
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

        // === Missing helpers restored ===
        private static string BuildPartyInfo(Outlook.MailItem mi)
        {
            try
            {
                var sb = new StringBuilder();
                var from = ActivityPartyBuilder.GetSenderSmtp(mi);
                if (!string.IsNullOrWhiteSpace(from)) sb.Append("from:").Append(from).Append(";");
                AppendRecipients(sb, mi, Outlook.OlMailRecipientType.olTo, "to");
                AppendRecipients(sb, mi, Outlook.OlMailRecipientType.olCC, "cc");
                AppendRecipients(sb, mi, Outlook.OlMailRecipientType.olBCC, "bcc");
                return sb.ToString();
            }
            catch { return null; }
        }
        private static void AppendRecipients(StringBuilder sb, Outlook.MailItem mi, Outlook.OlMailRecipientType type, string label)
        {
            try
            {
                bool first = true;
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
                    if (first) { sb.Append(label).Append(":"); first = false; } else { sb.Append(","); }
                    sb.Append(smtp);
                }
                if (!first) sb.Append(";");
            }
            catch { }
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
                        .OfType<StatusOptionMetadata>()
                        .FirstOrDefault(o => o.State == 1);
                    if (opt != null && opt.Value.HasValue) statusForCompleted = opt.Value.Value;
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
    }
}
