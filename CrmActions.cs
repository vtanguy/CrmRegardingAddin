using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;

// VT V2V2

namespace CrmRegardingAddin
{
    internal static class CrmActions
    {
        public static void SetRegarding(IOrganizationService org, EntityReference regarding, Outlook.MailItem mi)
        {
            if (org == null || mi == null) return;

            try
            {
                string mySmtp = GetCurrentUserSmtp(org);

                string fromSmtp = GetSenderSmtpStrict(mi);
                var toList  = new List<string>();
                var ccList  = new List<string>();
                var bccList = new List<string>();
                GetRecipients(mi, toList, ccList, bccList);

                bool isOutgoing = !string.IsNullOrEmpty(mySmtp) && !string.IsNullOrEmpty(fromSmtp) &&
                                  string.Equals(mySmtp, fromSmtp, StringComparison.OrdinalIgnoreCase);

                Guid meId = GetCurrentUserId(org);

                List<Entity> fromParties = BuildFromParties(org, fromSmtp, isOutgoing, meId);
                List<Entity> toParties   = BuildRecipients(org, toList,  2);
                List<Entity> ccParties   = BuildRecipients(org, ccList,  3);
                List<Entity> bccParties  = BuildRecipients(org, bccList, 4);

                string internetId = null;
                try { internetId = MailUtil.GetInternetMessageId(mi); } catch { }
                Entity existing = null;
                try { if (!string.IsNullOrEmpty(internetId)) existing = MailUtil.FindCrmEmailByMessageId(org, internetId); } catch { }

                Guid activityId;

                if (existing == null)
                {
                    var email = new Entity("email");
                    email["subject"]       = mi.Subject ?? "";
                    try { email["description"] = string.IsNullOrEmpty(mi.HTMLBody) ? (mi.Body ?? "") : mi.HTMLBody; } catch { email["description"] = mi.Body ?? ""; }
                    email["directioncode"] = isOutgoing;
                    if (!string.IsNullOrEmpty(internetId)) email["messageid"] = internetId;
                    if (regarding != null) email["regardingobjectid"] = regarding;
                    if (meId != Guid.Empty) email["ownerid"] = new EntityReference("systemuser", meId);

                    email["from"] = new EntityCollection(fromParties) { EntityName = "activityparty" };
                    if (toParties.Count  > 0) email["to"]  = new EntityCollection(toParties)  { EntityName = "activityparty" };
                    if (ccParties.Count  > 0) email["cc"]  = new EntityCollection(ccParties)  { EntityName = "activityparty" };
                    if (bccParties.Count > 0) email["bcc"] = new EntityCollection(bccParties) { EntityName = "activityparty" };

                    if (!isOutgoing && mi.ReceivedTime != DateTime.MinValue) email["actualstart"] = mi.ReceivedTime;
                    if ( isOutgoing && mi.SentOn       != DateTime.MinValue) email["actualend"]   = mi.SentOn;

                    activityId = org.Create(email);
                    CloseEmailAsCompleted(org, activityId);
                }
                else
                {
                    var upd = new Entity("email") { Id = existing.Id };
                    if (regarding != null) upd["regardingobjectid"] = regarding;

                    upd["from"] = new EntityCollection(fromParties) { EntityName = "activityparty" };
                    if (toParties.Count  > 0) upd["to"]  = new EntityCollection(toParties)  { EntityName = "activityparty" };
                    if (ccParties.Count  > 0) upd["cc"]  = new EntityCollection(ccParties)  { EntityName = "activityparty" };
                    if (bccParties.Count > 0) upd["bcc"] = new EntityCollection(bccParties) { EntityName = "activityparty" };

                    org.Update(upd);
                    CloseEmailAsCompleted(org, existing.Id);
                    activityId = existing.Id;
                }

                try
                {
                    var allRecipients = new List<string>();
                    allRecipients.AddRange(toList); allRecipients.AddRange(ccList); allRecipients.AddRange(bccList);

                    string regardingDisplay = (regarding != null && !string.IsNullOrEmpty(regarding.Name)) ? regarding.Name : "";
                    bool isIncoming = !isOutgoing;

                    allRecipients.AddRange(toList);
                    allRecipients.AddRange(ccList);
                    allRecipients.AddRange(bccList);

                    // ➜ NOUVEAU : récupérer l’OrgId
                    var orgId = GetOrganizationId(org);

                    // ➜ NOUVEAU : passer orgId et regardingDisplay à l’Interop
                    CrmMapiInterop.ApplyMsCompatForMail(
                        mi,
                        "", // crmid libre si tu ne veux pas le pousser
                        (regarding != null) ? regarding.Id : Guid.Empty,
                        regardingDisplay,
                        mySmtp,
                        (meId != Guid.Empty) ? (Guid?)meId : null,
                        fromSmtp,
                        allRecipients,
                        isIncoming,
                        orgId               // <— nouveau paramètre
                    );

                }
                catch { }

                try
                {
                    var insp = mi.GetInspector;
                    if (insp != null)
                        Globals.ThisAddIn.CreatePaneForMailIfLinked(insp, mi);
                } catch { }
            }
            catch (Exception)
            {
            }
        }

        public static void SetRegarding(IOrganizationService org, EntityReference regarding, Outlook.AppointmentItem appt)
        {
            if (org == null || appt == null) return;

            try
            {
                string globalId = null;
                try { globalId = appt.GlobalAppointmentID; } catch { }

                var existing = FindCrmAppointmentByGlobalId(org, globalId);

                var apptEntity = (existing == null)
                    ? new Entity("appointment")
                    : new Entity("appointment") { Id = existing.Id };

                if (!string.IsNullOrEmpty(globalId))
                    apptEntity["globalobjectid"] = globalId;

                if (existing == null)
                {
                    apptEntity["subject"] = appt.Subject ?? "";
                    var body = appt.Body;
                    if (!string.IsNullOrEmpty(body)) apptEntity["description"] = body;
                }

                if (appt.Start != DateTime.MinValue) apptEntity["scheduledstart"] = appt.Start;
                if (appt.End   != DateTime.MinValue) apptEntity["scheduledend"]   = appt.End;

                if (regarding != null) apptEntity["regardingobjectid"] = regarding;

                var meId = GetCurrentUserId(org);
                if (meId != Guid.Empty)
                {
                    var organizer = new Entity("activityparty");
                    organizer["participationtypemask"] = new OptionSetValue(7); // organizer
                    organizer["partyid"] = new EntityReference("systemuser", meId);
                    apptEntity["organizer"] = new EntityCollection(new List<Entity> { organizer }) { EntityName = "activityparty" };
                }

                Guid activityId;
                if (existing == null)
                {
                    activityId = org.Create(apptEntity);
                }
                else
                {
                    org.Update(apptEntity);
                    activityId = existing.Id;
                }

                try
                {
                    var mySmtp = GetCurrentUserSmtp(org);
                    var myId   = GetCurrentUserId(org);

                    string regardingDisplay = (regarding != null && !string.IsNullOrEmpty(regarding.Name)) ? regarding.Name : "";

                    CrmMapiInterop.ApplyMsCompatForAppointment(
                        appt,
                        (regarding != null) ? regarding.Id : Guid.Empty,
                        regardingDisplay,
                        mySmtp,
                        (myId != Guid.Empty) ? (Guid?)myId : null
                    );
                }
                catch { }

                try
                {
                    var insp = appt.GetInspector;
                    if (insp != null)
                        Globals.ThisAddIn.CreatePaneForAppointmentIfLinked(insp, appt);
                } catch { }
            }
            catch (Exception)
            {
            }
        }

        public static void UnlinkOrDeleteCrmEmail(IOrganizationService org, Outlook.MailItem mi, bool deleteInCrm)
        {
            if (org == null || mi == null) return;

            try
            {
                string msgId = null;
                try { msgId = MailUtil.GetInternetMessageId(mi); } catch { }
                Entity existing = null;
                try { if (!string.IsNullOrEmpty(msgId)) existing = MailUtil.FindCrmEmailByMessageId(org, msgId); } catch { }

                if (existing != null && deleteInCrm)
                {
                    try { org.Delete("email", existing.Id); } catch { }
                }

                try { CrmMapiInterop.RemoveMsCompatFromMail(mi); } catch { }

                try
                {
                    var insp = mi.GetInspector;
                    if (insp != null)
                        Globals.ThisAddIn.CreatePaneForMailIfLinked(insp, mi);
                } catch { }
            }
            catch (Exception)
            {
            }
        }

        public static void UnlinkOrDeleteCrmAppointment(IOrganizationService org, Outlook.AppointmentItem appt, bool deleteInCrm)
        {
            if (org == null || appt == null) return;

            try
            {
                string globalId = null; try { globalId = appt.GlobalAppointmentID; } catch { }
                var existing = FindCrmAppointmentByGlobalId(org, globalId, false);
                if (existing != null && deleteInCrm)
                {
                    try { org.Delete("appointment", existing.Id); } catch { }
                }

                try { CrmMapiInterop.RemoveMsCompatFromAppointment(appt); } catch { }

                try
                {
                    var insp = appt.GetInspector;
                    if (insp != null)
                        Globals.ThisAddIn.CreatePaneForAppointmentIfLinked(insp, appt);
                } catch { }
            }
            catch (Exception)
            {
            }
        }

        public static void UnlinkOrDeleteCrmAppointment(IOrganizationService org, Outlook.AppointmentItem appt)
        {
            UnlinkOrDeleteCrmAppointment(org, appt, false);
        }

        private static List<Entity> BuildFromParties(IOrganizationService org, string fromSmtp, bool isOutgoing, Guid meId)
        {
            var list = new List<Entity>();
            if (isOutgoing && meId != Guid.Empty)
            {
                list.Add(BuildActivityPartySystemUser(meId, 1)); // From = 1
            }
            else
            {
                var ap = new Entity("activityparty");
                ap["participationtypemask"] = new OptionSetValue(1);
                ap["addressused"] = string.IsNullOrEmpty(fromSmtp) ? "(unknown)" : fromSmtp;
                list.Add(ap);
            }
            return list;
        }

        private static List<Entity> BuildRecipients(IOrganizationService org, List<string> emails, int mask)
        {
            var bags = new List<Entity>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in emails)
            {
                if (string.IsNullOrWhiteSpace(s)) continue;
                if (seen.Contains(s)) continue;
                seen.Add(s);

                bags.Add(BuildActivityPartyFromEmail(org, s, mask));
            }
            return bags;
        }

        private static Entity BuildActivityPartyFromEmail(IOrganizationService org, string email, int participationTypeMask)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                var apx = new Entity("activityparty");
                apx["participationtypemask"] = new OptionSetValue(participationTypeMask);
                apx["addressused"] = "(unknown)";
                return apx;
            }

            var er = TryResolveByEmail(org, "contact", "emailaddress1", email, "fullname");
            if (er == null) er = TryResolveByEmail(org, "account", "emailaddress1", email, "name");
            if (er == null) er = TryResolveByEmail(org, "systemuser", "internalemailaddress", email, "fullname");

            var ap = new Entity("activityparty");
            ap["participationtypemask"] = new OptionSetValue(participationTypeMask);
            if (er != null) ap["partyid"] = er; else ap["addressused"] = email;
            return ap;
        }

        private static EntityReference TryResolveByEmail(IOrganizationService org, string entity, string emailField, string email, string nameField)
        {
            try
            {
                var q = new QueryExpression(entity)
                {
                    ColumnSet = new ColumnSet(nameField),
                    TopCount = 1,
                    NoLock = true
                };
                q.Criteria.AddCondition(emailField, ConditionOperator.Equal, email);
                var res = org.RetrieveMultiple(q).Entities;
                if (res.Count > 0)
                {
                    var e = res[0];
                    var er = new EntityReference(entity, e.Id);
                    var name = e.GetAttributeValue<string>(nameField);
                    if (!string.IsNullOrEmpty(name)) er.Name = name;
                    return er;
                }
            }
            catch { }
            return null;
        }

        private static Entity BuildActivityPartySystemUser(Guid systemUserId, int participationTypeMask)
        {
            var ap = new Entity("activityparty");
            ap["participationtypemask"] = new OptionSetValue(participationTypeMask);
            ap["partyid"] = new EntityReference("systemuser", systemUserId);
            return ap;
        }

        private static Guid GetCurrentUserId(IOrganizationService org)
        {
            try
            {
                var req = new OrganizationRequest("WhoAmI");
                var resp = org.Execute(req) as OrganizationResponse;
                if (resp != null && resp.Results != null && resp.Results.Contains("UserId"))
                    return (Guid)resp.Results["UserId"];
            }
            catch { }
            return Guid.Empty;
        }

        private static string GetCurrentUserSmtp(IOrganizationService org)
        {
            try
            {
                var me = GetCurrentUserId(org);
                if (me == Guid.Empty) return null;
                var usr = org.Retrieve("systemuser", me, new ColumnSet("internalemailaddress"));
                return usr.GetAttributeValue<string>("internalemailaddress");
            }
            catch { return null; }
        }

        private static void GetRecipients(Outlook.MailItem mi, List<string> to, List<string> cc, List<string> bcc)
        {
            try
            {
                var recips = mi.Recipients;
                if (recips != null)
                {
                    foreach (Outlook.Recipient r in recips)
                    {
                        if (r == null) continue;
                        try { if (!r.Resolved) r.Resolve(); } catch { }

                        var smtp = AddressEntryToSmtp(r.AddressEntry, r.Address);
                        if (string.IsNullOrEmpty(smtp)) continue;

                        switch (r.Type)
                        {
                            case (int)Outlook.OlMailRecipientType.olTo:  to.Add(smtp);  break;
                            case (int)Outlook.OlMailRecipientType.olCC:  cc.Add(smtp);  break;
                            case (int)Outlook.OlMailRecipientType.olBCC: bcc.Add(smtp); break;
                        }
                    }
                }
            }
            catch { }

            if (to.Count == 0 && !string.IsNullOrEmpty(mi.To))   to.AddRange(ParseAddressList(mi, mi.To));
            if (cc.Count == 0 && !string.IsNullOrEmpty(mi.CC))   cc.AddRange(ParseAddressList(mi, mi.CC));
            if (bcc.Count == 0 && !string.IsNullOrEmpty(mi.BCC)) bcc.AddRange(ParseAddressList(mi, mi.BCC));
        }

        private static IEnumerable<string> ParseAddressList(Outlook.MailItem mi, string disp)
        {
            var app = mi.Application as Outlook.Application;
            var parts = (disp ?? "").Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p0 in parts)
            {
                var p = (p0 ?? "").Trim().Trim('\"');
                if (string.IsNullOrEmpty(p)) continue;

                string smtp = null;
                try
                {
                    var temp = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    var r = temp.Recipients.Add(p);
                    r.Resolve();
                    smtp = AddressEntryToSmtp(r.AddressEntry, r.Address);
                    try { temp.Close(Outlook.OlInspectorClose.olDiscard); } catch { }
                }
                catch { }

                yield return string.IsNullOrEmpty(smtp) ? p : smtp;
            }
        }

        private static string AddressEntryToSmtp(Outlook.AddressEntry ae, string fallback)
        {
            try
            {
                if (ae != null && ae.Type != null && ae.Type.Equals("EX", StringComparison.OrdinalIgnoreCase))
                {
                    var exUser = ae.GetExchangeUser();
                    if (exUser != null && !string.IsNullOrEmpty(exUser.PrimarySmtpAddress))
                        return exUser.PrimarySmtpAddress;

                    var exDL = ae.GetExchangeDistributionList();
                    if (exDL != null && !string.IsNullOrEmpty(exDL.PrimarySmtpAddress))
                        return exDL.PrimarySmtpAddress;
                }
                if (ae != null && !string.IsNullOrEmpty(ae.Address))
                    return ae.Address;
            }
            catch { }
            return fallback;
        }

        private static string GetSenderSmtpStrict(Outlook.MailItem mi)
        {
            try
            {
                var pa = mi.PropertyAccessor;
                const string HDR = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
                var headers = pa.GetProperty(HDR) as string;
                if (!string.IsNullOrEmpty(headers))
                {
                    var lines = headers.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    foreach (var line in lines)
                    {
                        if (!line.StartsWith("From:", StringComparison.OrdinalIgnoreCase)) continue;
                        var m = System.Text.RegularExpressions.Regex.Match(line, @"<([^>]+)>");
                        if (m.Success) return m.Groups[1].Value.Trim();
                        m = System.Text.RegularExpressions.Regex.Match(
                                line,
                                @"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        if (m.Success) return m.Value.Trim();
                    }
                }
            }
            catch { }

            try
            {
                var smtp = mi.SenderEmailAddress;
                if (!string.IsNullOrEmpty(smtp)) return smtp;
            }
            catch { }

            try
            {
                var sender = mi.Sender;
                if (sender != null)
                    return AddressEntryToSmtp(sender, null);
            }
            catch { }

            return null;
        }

        private static void CloseEmailAsCompleted(IOrganizationService org, Guid emailId)
        {
            try
            {
                var req = new OrganizationRequest("SetState")
                {
                    ["EntityMoniker"] = new EntityReference("email", emailId),
                    ["State"]  = new OptionSetValue(1),
                    ["Status"] = new OptionSetValue(-1)
                };
                org.Execute(req);
            }
            catch { }
        }

        public static Entity FindCrmAppointmentByGlobalObjectId(IOrganizationService org, string id)
        {
            return FindCrmAppointmentByGlobalId(org, id, true);
        }

        public static Entity FindCrmAppointmentByGlobalId(IOrganizationService org, string globalId, bool onlyIfLinked = true)
        {
            if (org == null || string.IsNullOrEmpty(globalId)) return null;

            var qe = new QueryExpression("appointment")
            {
                ColumnSet = new ColumnSet("activityid", "subject", "globalobjectid", "regardingobjectid")
            };
            qe.Criteria.AddCondition("globalobjectid", ConditionOperator.Equal, globalId);
            if (onlyIfLinked)
                qe.Criteria.AddCondition("regardingobjectid", ConditionOperator.NotNull);

            var res = org.RetrieveMultiple(qe);
            return (res != null && res.Entities != null && res.Entities.Count > 0) ? res.Entities[0] : null;
        }
        private static Guid GetOrganizationId(IOrganizationService org)
        {
            try
            {
                var q = new Microsoft.Xrm.Sdk.Query.QueryExpression("organization")
                {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("organizationid"),
                    TopCount = 1,
                    NoLock = true
                };
                var r = org.RetrieveMultiple(q);
                if (r != null && r.Entities.Count > 0)
                    return r.Entities[0].Id;
            }
            catch { }
            return Guid.Empty;
        }
    }
}
