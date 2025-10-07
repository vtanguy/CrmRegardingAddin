
﻿using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

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

                var resolvedPartyCache = new Dictionary<string, EntityReference>(StringComparer.OrdinalIgnoreCase);

                List<Entity> fromParties = BuildFromParties(org, fromSmtp, isOutgoing, meId);
                List<Entity> toParties   = BuildRecipients(org, toList,  2, resolvedPartyCache);
                List<Entity> ccParties   = BuildRecipients(org, ccList,  3, resolvedPartyCache);
                List<Entity> bccParties  = BuildRecipients(org, bccList, 4, resolvedPartyCache);

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
                    // Destinataires consolidés, sans double ajout
                    var allRecipients = new List<string>();
                    allRecipients.AddRange(toList);
                    allRecipients.AddRange(ccList);
                    allRecipients.AddRange(bccList);

                    // Regarding display : si vide dans l'EntityReference, on tente un retrieve pour peupler
                    string regardingDisplay = (regarding != null && !string.IsNullOrEmpty(regarding.Name))
                                              ? regarding.Name
                                              : TryGetRegardingDisplay(org, regarding);

                    bool isIncoming = !isOutgoing;

                    // Code objet attendu par l'addin MS (ex: contact=2, account=1, lead=4, opportunity=3...)
                    var regardingTypeCode = MapEntityLogicalNameToObjectTypeCode(regarding != null ? regarding.LogicalName : null);

                    // Id d'organisation pour crmorgid
                    var orgId = GetOrganizationId(org);

                    var compatRecipients = BuildCompatPartyMembers(org, allRecipients, resolvedPartyCache);
                    var fromMember = isOutgoing
                        ? BuildSystemUserCompatPartyMember(mySmtp, meId)
                        : BuildCompatPartyMember(org, fromSmtp, resolvedPartyCache);

                    CrmMapiInterop.ApplyMsCompatForMail(
                        mi,
                        regardingTypeCode,
                        (regarding != null) ? regarding.Id : Guid.Empty,
                        regardingDisplay ?? "",
                        mySmtp,
                        (meId != Guid.Empty) ? (Guid?)meId : null,
                        fromMember,
                        compatRecipients,
                        isIncoming,
                        orgId
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

            Logger.Info("[Appt] SetRegarding: subject='" + (appt.Subject ?? "") + "'");

            // 1) Clé unique Outlook
            string globalId = null;
            try { globalId = appt.GlobalAppointmentID; } catch { }
            Logger.Info("[Appt] GlobalAppointmentID=" + (globalId ?? "(null)"));

            // 2) Recherche CRM par globalobjectid (sans forcer 'regarding' pour aussi retrouver ceux non liés)
            var existing = FindCrmAppointmentByGlobalId(org, globalId, onlyIfLinked: false);
            Logger.Info(existing == null ? "[Appt] Not found in CRM" : "[Appt] Found CRM appt: " + existing.Id);

            // 3) Contexte utilisateur/org
            var meId = GetCurrentUserId(org);
            var ownerSmtp = GetCurrentUserSmtp(org);
            var orgId = GetOrganizationId(org);
            string regardingDisplay = TryGetRegardingDisplay(org, regarding);

            // 4) Créer si absent (évite le doublon SSS car même 'globalobjectid')
            Guid crmApptId = Guid.Empty;
            if (existing == null)
            {
                var e = new Entity("appointment");
                if (!string.IsNullOrEmpty(globalId)) e["globalobjectid"] = globalId;
                e["subject"] = appt.Subject ?? "";
                try { e["scheduledstart"] = appt.Start; } catch { }
                try { e["scheduledend"] = appt.End; } catch { }
                if (regarding != null) e["regardingobjectid"] = regarding;
                if (meId != Guid.Empty) e["ownerid"] = new EntityReference("systemuser", meId);

                try
                {
                    crmApptId = org.Create(e);
                    Logger.Info("[Appt] Created CRM appt: " + crmApptId);
                }
                catch (Exception ex)
                {
                    Logger.Info("[Appt] Create EX: " + ex);
                }
            }
            else
            {
                crmApptId = existing.Id;
                // 5) Mise à jour du regarding si demandé
                if (regarding != null)
                {
                    var upd = new Entity("appointment") { Id = existing.Id };
                    upd["regardingobjectid"] = regarding;
                    try { org.Update(upd); Logger.Info("[Appt] Updated CRM appt: " + existing.Id); }
                    catch (Exception ex) { Logger.Info("[Appt] Update EX: " + ex); }
                }
            }

            // 6) Tag UDF côté Outlook (fallback UserProperties si MAPI bloqué) + crmid s’il est connu
            bool tagOk = OutlookApptTagger.TagAppointmentLinkWithDiagnostics(
                appt,
                crmApptId == Guid.Empty ? (Guid?)null : crmApptId,
                regarding != null ? regarding.Id : Guid.Empty,
                regardingDisplay ?? "",
                MapEntityLogicalNameToObjectTypeCode(regarding != null ? regarding.LogicalName : null),
                orgId,
                ownerSmtp,
                meId);

            // 7) Feedback visuel dans le sujet (optionnel, simple)
            try
            {
                var suffix = tagOk ? "[UDF OK]" : "[UDF ERR]"; // ne change rien côté CRM
                if ((appt.Subject ?? "").IndexOf(suffix, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    appt.Subject = (appt.Subject ?? "") + " " + suffix;
                    appt.Save();
                    Logger.Info("[Appt] Subject tagged with " + suffix);
                }
            }
            catch (Exception ex) { Logger.Info("[Appt] Subject tag EX: " + ex.Message); }
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
                try { Logger.Info("[APPT] Unlink/Delete start GlobalAppointmentID=" + (globalId ?? "<null>")); } catch {}
                var existing = FindCrmAppointmentByGlobalId(org, globalId, false);
                if (existing != null && deleteInCrm)
                {
                    try { org.Delete("appointment", existing.Id); Logger.Info("[APPT] Deleted CRM appointment ActivityId=" + existing.Id.ToString()); } catch { }
                }

                try { CrmMapiInterop.RemoveMsCompatFromAppointment(appt); Logger.Info("[APPT] Removed MS compat named props from Outlook item"); } catch { }

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

        private static List<Entity> BuildRecipients(
            IOrganizationService org,
            List<string> emails,
            int mask,
            Dictionary<string, EntityReference> resolvedPartyCache)
        {
            var bags = new List<Entity>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in emails)
            {
                var normalized = (s ?? "").Trim();
                if (normalized.Length == 0) continue;
                if (!seen.Add(normalized)) continue;

                bags.Add(BuildActivityPartyFromEmail(org, normalized, mask, resolvedPartyCache));
            }
            return bags;
        }

        private static Entity BuildActivityPartyFromEmail(
            IOrganizationService org,
            string email,
            int participationTypeMask,
            Dictionary<string, EntityReference> resolvedPartyCache)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                var apx = new Entity("activityparty");
                apx["participationtypemask"] = new OptionSetValue(participationTypeMask);
                apx["addressused"] = "(unknown)";
                return apx;
            }

            var er = ResolvePartyByEmail(org, email, resolvedPartyCache);

            var ap = new Entity("activityparty");
            ap["participationtypemask"] = new OptionSetValue(participationTypeMask);
            if (er != null) ap["partyid"] = er; else ap["addressused"] = email;
            return ap;
        }

        private static EntityReference ResolvePartyByEmail(
            IOrganizationService org,
            string email,
            Dictionary<string, EntityReference> resolvedPartyCache)
        {
            if (string.IsNullOrWhiteSpace(email)) return null;

            EntityReference cached;
            if (resolvedPartyCache != null && resolvedPartyCache.TryGetValue(email, out cached))
                return cached;

            var er = TryResolveByEmail(org, "contact", "emailaddress1", email, "fullname");
            if (er == null) er = TryResolveByEmail(org, "account", "emailaddress1", email, "name");
            if (er == null) er = TryResolveByEmail(org, "systemuser", "internalemailaddress", email, "fullname");

            if (resolvedPartyCache != null)
                resolvedPartyCache[email] = er;

            return er;
        }

        private static IEnumerable<CrmMapiInterop.CrmPartyMember> BuildCompatPartyMembers(
            IOrganizationService org,
            IEnumerable<string> emails,
            Dictionary<string, EntityReference> resolvedPartyCache)
        {
            var list = new List<CrmMapiInterop.CrmPartyMember>();
            if (emails == null) return list;

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in emails)
            {
                var email = (raw ?? "").Trim();
                if (email.Length == 0) continue;
                if (!seen.Add(email)) continue;

                var member = BuildCompatPartyMember(org, email, resolvedPartyCache);
                if (member != null) list.Add(member);
            }

            return list;
        }

        private static CrmMapiInterop.CrmPartyMember BuildCompatPartyMember(
            IOrganizationService org,
            string email,
            Dictionary<string, EntityReference> resolvedPartyCache)
        {
            if (string.IsNullOrWhiteSpace(email)) return null;

            var normalized = email.Trim();
            var member = new CrmMapiInterop.CrmPartyMember
            {
                Email = normalized,
                Name = normalized,
                TypeCode = -1
            };

            var er = ResolvePartyByEmail(org, normalized, resolvedPartyCache);
            if (er != null)
            {
                member.PartyId = er.Id;
                if (!string.IsNullOrEmpty(er.Name))
                    member.Name = er.Name;

                int mapped;
                var mappedString = MapEntityLogicalNameToObjectTypeCode(er.LogicalName);
                if (!string.IsNullOrEmpty(mappedString) && int.TryParse(mappedString, out mapped))
                    member.TypeCode = mapped;
            }

            return member;
        }

        private static CrmMapiInterop.CrmPartyMember BuildSystemUserCompatPartyMember(string email, Guid systemUserId)
        {
            if (string.IsNullOrWhiteSpace(email)) return null;
            return new CrmMapiInterop.CrmPartyMember
            {
                Email = email,
                Name = email,
                PartyId = (systemUserId != Guid.Empty) ? (Guid?)systemUserId : null,
                TypeCode = 8
            };
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
                if (ae == null) return fallback;

                // 1) PR_SMTP_ADDRESS (0x39FE001E)
                try
                {
                    var pa = ae.PropertyAccessor;
                    if (pa != null)
                    {
                        const string PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var val = pa.GetProperty(PR_SMTP) as string;
                        if (!string.IsNullOrEmpty(val))
                            return val;
                    }
                }
                catch { /* ignore */ }

                // 2) Exchange user / DL PrimarySmtpAddress
                if (ae.Type != null && ae.Type.Equals("EX", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var exUser = ae.GetExchangeUser();
                        if (exUser != null && !string.IsNullOrEmpty(exUser.PrimarySmtpAddress))
                            return exUser.PrimarySmtpAddress;
                    }
                    catch { }
                    try
                    {
                        var exDL = ae.GetExchangeDistributionList();
                        if (exDL != null && !string.IsNullOrEmpty(exDL.PrimarySmtpAddress))
                            return exDL.PrimarySmtpAddress;
                    }
                    catch { }
                }

                // 3) SMTP explicite
                if (ae.Type != null && ae.Type.Equals("SMTP", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrEmpty(ae.Address))
                {
                    return ae.Address;
                }

                // 4) Ne jamais retourner un DN X.500
                if (!string.IsNullOrEmpty(ae.Address) && ae.Address.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                    return fallback;

                // 5) Dernier recours : certains providers stockent la SMTP ici
                if (!string.IsNullOrEmpty(ae.Address))
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
                if (!string.IsNullOrEmpty(smtp) && !smtp.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                    return smtp;
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

        try
        {
            var res = org.RetrieveMultiple(qe);
            var result = (res != null && res.Entities != null && res.Entities.Count > 0) ? res.Entities[0] : null;
            Logger.Info(result == null ? "[Appt] RetrieveMultiple: 0" : "[Appt] RetrieveMultiple: found " + result.Id);
            return result;
        }
        catch (Exception ex) { Logger.Info("[Appt] RetrieveMultiple EX: " + ex.Message); return null; }
    }
    private static Guid GetOrganizationId(IOrganizationService org)
        {
            try
            {
                var q = new QueryExpression("organization")
                {
                    ColumnSet = new ColumnSet("organizationid"),
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

        private static string TryGetRegardingDisplay(IOrganizationService org, EntityReference regarding)
        {
            if (org == null || regarding == null) return null;
            if (!string.IsNullOrEmpty(regarding.Name)) return regarding.Name;
            try
            {
                var e = org.Retrieve(regarding.LogicalName, regarding.Id, new ColumnSet("name", "fullname", "subject"));
                return e.GetAttributeValue<string>("name") ?? e.GetAttributeValue<string>("fullname") ?? e.GetAttributeValue<string>("subject");
            }
            catch { return null; }
        }

        private static string MapEntityLogicalNameToObjectTypeCode(string logicalName)
        {
            if (string.IsNullOrEmpty(logicalName)) return "";
            switch (logicalName)
            {
                case "account": return "1";
                case "contact": return "2";
                case "opportunity": return "3";
                case "lead": return "4";
                case "systemuser": return "8";
                case "incident": return "112";
                default: return "";
            }
        }
        private static class OutlookApptTagger
        {
            private const string Pfx = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

            private static bool TrySetProperty(Outlook.PropertyAccessor pa, string name, object value, Outlook.UserProperties userProps)
            {
                try { pa.SetProperty(Pfx + name, value); Logger.Info("[UDF] Set OK: " + name); return true; }
                catch (Exception ex)
                {
                    Logger.Info("[UDF] Set FAIL: " + name + " EX=" + ex.Message + " -> fallback UserProperty");
                    try
                    {
                        var up = userProps[name] ?? userProps.Add(name, Outlook.OlUserPropertyType.olText);
                        up.Value = value == null ? "" : value.ToString();
                        Logger.Info("[UDF] Fallback OK: " + name);
                        return true;
                    }
                    catch (Exception ex2)
                    {
                        Logger.Info("[UDF] Fallback FAIL: " + name + " EX=" + ex2.Message);
                        return false;
                    }
                }
            }

            public static bool TagAppointmentLinkWithDiagnostics(
                Outlook.AppointmentItem appt,
                Guid? crmApptId,
                Guid regardingId,
                string regardingName,
                string regardingOtc,
                Guid orgId,
                string ownerSmtp,
                Guid ownerUserId)
            {
                bool ok = true;
                try
                {
                    var pa = appt.PropertyAccessor;
                    var ups = appt.UserProperties;
                    if (crmApptId.HasValue) ok &= TrySetProperty(pa, "crmid", crmApptId.Value.ToString("B"), ups);
                    ok &= TrySetProperty(pa, "crmlinkstate", 1.0, ups);
                    ok &= TrySetProperty(pa, "crmorgid", orgId == Guid.Empty ? "" : orgId.ToString("B"), ups);
                    ok &= TrySetProperty(pa, "crmownersmtp", ownerSmtp ?? "", ups);
                    ok &= TrySetProperty(pa, "crmownersystemuserid", ownerUserId == Guid.Empty ? "" : ownerUserId.ToString("B"), ups);
                    ok &= TrySetProperty(pa, "crmregardingobjectid", regardingId == Guid.Empty ? "" : regardingId.ToString("B"), ups);
                    ok &= TrySetProperty(pa, "crmregardingobject", regardingName ?? "", ups);
                    ok &= TrySetProperty(pa, "crmregardingobjecttypecode", regardingOtc ?? "", ups);
                    try { appt.Save(); } catch { }
                    Logger.Info("[Appt] Tag UDF done.");
                    return ok;
                }
                catch (Exception ex) { Logger.Info("[Appt] Tag UDF EX: " + ex); return false; }
            }
        }
    }
}
