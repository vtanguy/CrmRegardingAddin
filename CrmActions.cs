using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Windows.Forms;
//VT 11-10
namespace CrmRegardingAddin
{
    internal static class CrmActions
    {
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

                try { CrmMapiInterop.RemoveMsCompatFromMail(mi);
                try { mi.Save(); } catch { }
                } catch { }

                try
                {
                    var insp = mi.GetInspector;
                    if (insp != null)
                        Globals.ThisAddIn.CreatePaneForMailIfLinked(insp, mi);
                } catch { }
            }
            catch (Exception)
            {     }
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

                try { CrmMapiInterop.RemoveMsCompatFromAppointment(appt);
                try { appt.Save(); } catch { }

                } catch { }

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

        // === New link/unlink entry points ===

        /// <summary>
        /// Prompts for a CRM record and links the current MailItem to it.
        /// - Creates/updates the Email activity in CRM (messageid = InternetMessageId)
        /// - Stamps Outlook named properties exactly like the Microsoft add-in (subset: LinkState, RegardingId/Type, SssPromote, crmid/orgid when available)
        /// - Optionally marks the 'Tracked' category depending on App.config
        /// </summary>
        
        public static void CreateLinkForMail(IOrganizationService org, Outlook.MailItem mi)
        {
            if (org == null || mi == null) return;

            // 1) Suggest regarding (Contact) from FROM/TO
            EntityReference regarding = null;
            try
            {
                var suggestion = SuggestContactRegardingFromMail(org, mi);
                if (suggestion != null)
                {
                    var label = GetReadableName(org, suggestion);
                    var dr = MessageBox.Show(
                        string.Format("Contact trouvé : {0}.\r\nVoulez-vous lier cet e-mail à ce contact ?", label),
                        "CRM", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes) regarding = suggestion;
                }
            } catch { }

            // 2) Sinon, demander via la recherche standard
            if (regarding == null)
                regarding = PromptForRegarding(org);
            if (regarding == null) return;

            // 3) Find or create CRM Email by InternetMessageId
            var internetId = MailUtil.GetInternetMessageId(mi);
            if (string.IsNullOrEmpty(internetId))
                throw new InvalidOperationException("Impossible de récupérer l'InternetMessageId.");

            Entity crmEmail = MailUtil.FindCrmEmailByMessageId(org, internetId);
            Guid emailId;
            bool outgoing = IsOutgoing(mi);

            if (crmEmail == null)
            {
                crmEmail = new Entity("email");
                crmEmail["subject"] = Safe(() => mi.Subject) ?? "(sans objet)";
                crmEmail["description"] = Safe(() => mi.HTMLBody) ?? Safe(() => mi.Body);
                crmEmail["messageid"] = internetId;
                crmEmail["regardingobjectid"] = regarding;
                crmEmail["directioncode"] = outgoing;

                // Parties (with participationtypemask + partyid resolution)
                crmEmail["from"] = new Microsoft.Xrm.Sdk.EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olOriginator, org));
                crmEmail["to"]   = new Microsoft.Xrm.Sdk.EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olTo,         org));
                crmEmail["cc"]   = new Microsoft.Xrm.Sdk.EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olCC,         org));
                crmEmail["bcc"]  = new Microsoft.Xrm.Sdk.EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olBCC,        org));

                emailId = org.Create(crmEmail);
            }
            else
            {
                emailId = crmEmail.Id;
                var upd = new Entity("email");
                upd.Id = emailId;
                upd["regardingobjectid"] = regarding;

                // Backfill parties if missing
                if (IsEmailDraft(org, crmEmail))
                {
                    TryBackfillPartiesIfEmpty(org, crmEmail, upd, mi);
                }

                // Ensure directioncode is set
                if (!crmEmail.Contains("directioncode"))
                    upd["directioncode"] = outgoing;

                org.Update(upd);
            }

            // 3) Fermer l'email (Completed: Sent/Received)
            MarkEmailCompleted(org, emailId, outgoing);

            // 4) Stamp Outlook
            var orgId = GetOrgId(org);
            CrmMapiInterop.SetCrmLinkPropsForMail(mi, regarding.Id, GetTypeCode(regarding), emailId, orgId);

            RecentLinks.Remember(regarding);

            // 5) Optional Outlook category and sync ping
            OutlookUtil.TryMarkTrackedCategory(mi);
            MsCrmCleanup.TriggerSyncIfPossible();
        }
        
        private static EntityReference ResolveContactByEmail(IOrganizationService org, string email)
        {
            if (org == null || string.IsNullOrWhiteSpace(email)) return null;
            return FindByAnyEmail(org, "contact", new[] { "emailaddress1", "emailaddress2", "emailaddress3" }, email);
        }

        /// <summary>
        /// Suggest Contact from FROM/TO addresses. Returns null if none or multiple.
        /// </summary>
        private static EntityReference SuggestContactRegardingFromMail(IOrganizationService org, Outlook.MailItem mi)
        {
            try
            {
                var from = GetSenderSmtp(mi);
                var erFrom = ResolveContactByEmail(org, from);
                if (erFrom != null) return erFrom;

                var distinct = new System.Collections.Generic.Dictionary<System.Guid, EntityReference>();
                var recips = mi?.Recipients;
                if (recips != null)
                {
                    foreach (Outlook.Recipient r in recips)
                    {
                        try
                        {
                            if (r == null || r.Type != (int)Outlook.OlMailRecipientType.olTo) continue;
                            var addr = GetSmtpAddress(r);
                            if (string.IsNullOrWhiteSpace(addr)) continue;
                            var er = ResolveContactByEmail(org, addr);
                            if (er != null && !distinct.ContainsKey(er.Id)) distinct[er.Id] = er;
                        } catch { }
                    }
                }
                if (distinct.Count == 1) return new System.Collections.Generic.List<EntityReference>(distinct.Values)[0];
            }
            catch { }
            return null;
        }
    
        public static void CreateLinkForAppointment(IOrganizationService org, Outlook.AppointmentItem appt)
        {
            if (org == null || appt == null) return;

            // 1) Prompt for Regarding
            EntityReference regarding = PromptForRegarding(org);
            if (regarding == null) return;

            // 2) Find or create CRM Appointment by GlobalAppointmentID
            string globalId = Safe(() => appt.GlobalAppointmentID);
            if (string.IsNullOrEmpty(globalId))
                throw new InvalidOperationException("Impossible de récupérer le GlobalAppointmentID.");

            Entity crmAppt = FindCrmAppointmentByGlobalId(org, globalId, false);
            Guid apptId;
            if (crmAppt == null)
            {
                crmAppt = new Entity("appointment");
                crmAppt["subject"] = Safe(() => appt.Subject) ?? "(sans objet)";
                crmAppt["description"] = Safe(() => appt.Body);
                var start = Safe(() => appt.Start);
                var endt  = Safe(() => appt.End);
                if (start != default(DateTime)) crmAppt["scheduledstart"] = start;
                if (endt  != default(DateTime)) crmAppt["scheduledend"]   = endt;
                crmAppt["globalobjectid"] = globalId;
                crmAppt["regardingobjectid"] = regarding;

                // Participants
                crmAppt["requiredattendees"] = new EntityCollection(BuildPartiesFromAppointment(appt, true, org));
                crmAppt["optionalattendees"] = new EntityCollection(BuildPartiesFromAppointment(appt, false, org));

                apptId = org.Create(crmAppt);
            }
            else
            {
                apptId = crmAppt.Id;
                var upd = new Entity("appointment") { Id = apptId };
                upd["regardingobjectid"] = regarding;
                TryBackfillApptPartiesIfEmpty(org, crmAppt, upd, appt);
                org.Update(upd);
            }

            // 3) Stamp Outlook props for appointment (objectTypeCode=4201, crmEntryID, etc. handled inside)
            var orgId = GetOrgId(org);
            CrmMapiInterop.SetCrmLinkPropsForAppointment(appt, regarding.Id, GetTypeCode(regarding), apptId, orgId, Safe(() => appt.EntryID));

            try { appt.Save(); } catch { }

            RecentLinks.Remember(regarding);

            // 4) Pane refresh
            try
            {
                var insp = appt.GetInspector;
                if (insp != null)
                    Globals.ThisAddIn.CreatePaneForAppointmentIfLinked(insp, appt);
            } catch { }
        }



        public static void UnlinkAppointment(IOrganizationService org, Outlook.AppointmentItem appt, bool keepCrmItem)
        {
            if (appt == null) return;
            CrmMapiInterop.RemoveCrmLinkProps(appt);
            if (!keepCrmItem)
            {
                try
                {
                    var goid = Safe(() => appt.GlobalAppointmentID);
                    var existing = FindCrmAppointmentByGlobalId(org, goid, false);
                    if (existing != null) org.Delete("appointment", existing.Id);
                } catch { }
            }
        }

        // === Helpers ===

        private static EntityReference PromptForRegarding(IOrganizationService org)
        {
            try
            {
                using (var dlg = new SearchDialog(org))
                {
                    return dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK ? dlg.SelectedReference : null;
                }
            }
            catch { return null; }
        }

        private static int GetTypeCode(EntityReference er)
        {
            // In CRM 8.x, EntityReference does not carry TypeCode; we must query once if needed.
            // To avoid metadata calls, rely on common built-ins when possible.
            // Fallback: derive from logical name using a minimal hard-map.
            switch ((er == null ? null : er.LogicalName) ?? "")
            {
                case "account": return 1;
                case "contact": return 2;
                case "opportunity": return 3;
                case "lead": return 4;
                case "systemuser": return 8;
                default: return 0; // unknown -> let add-in/SSS cope
            }
        }

        
        private static bool IsOutgoing(Outlook.MailItem mi)
        {
            try
            {
                var ns = Globals.ThisAddIn.Application.Session;
                var current = ns?.CurrentUser?.AddressEntry?.GetExchangeUser()?.PrimarySmtpAddress;
                var sender = mi.SenderEmailAddress;
                if (!string.IsNullOrEmpty(current) && !string.IsNullOrEmpty(sender))
                    return string.Equals(current, sender, StringComparison.OrdinalIgnoreCase);
            } catch { }
            return false;
        }

                
        private static bool IsEmailDraft(IOrganizationService org, Entity existing)
        {
            try
            {
                OptionSetValue sc = null;
                if (existing != null && existing.Attributes.ContainsKey("statuscode"))
                {
                    sc = existing.GetAttributeValue<OptionSetValue>("statuscode");
                }
                else if (existing != null)
                {
                    var e = org.Retrieve("email", existing.Id, new ColumnSet("statuscode"));
                    sc = e.GetAttributeValue<OptionSetValue>("statuscode");
                }
                // Email statuscode: 1 = Draft
                return (sc != null && sc.Value == 1);
            }
            catch
            {
                return false;
            }
        }
    private static void MarkEmailCompleted(IOrganizationService org, Guid emailId, bool outgoing)
        {
            if (org == null || emailId == Guid.Empty) return;
            try
            {
                // Transition to Completed with proper StatusReason (Sent/Received)
                var req = new OrganizationRequest("SetState");
                req["EntityMoniker"] = new EntityReference("email", emailId);
                req["State"] = new OptionSetValue(1);                 // Completed
                req["Status"] = new OptionSetValue(outgoing ? 3 : 4); // 3=Sent, 4=Received
                org.Execute(req);
            }
            catch
            {
                // Fallback: update statuscode if SetState is restricted
                try
                {
                    var upd = new Entity("email") { Id = emailId };
                    upd["statuscode"] = new OptionSetValue(outgoing ? 3 : 4);
                    org.Update(upd);
                }
                catch { }
            }
        }

        private static void TryBackfillPartiesIfEmpty(IOrganizationService org, Entity existing, Entity update, Outlook.MailItem mi)
        {
            try
            {
                bool hasFrom = existing.Contains("from") && existing.GetAttributeValue<EntityCollection>("from")?.Entities?.Count > 0;
                bool hasTo   = existing.Contains("to")   && existing.GetAttributeValue<EntityCollection>("to")?.Entities?.Count > 0;
                if (!hasFrom)
                    update["from"] = new EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olOriginator, org));
                if (!hasTo)
                {
                    update["to"]  = new EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olTo, org));
                    update["cc"]  = new EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olCC, org));
                    update["bcc"] = new EntityCollection(BuildPartiesFromMail(mi, Outlook.OlMailRecipientType.olBCC, org));
                }
            } catch { }
        }

        private static string GetSenderSmtp(Outlook.MailItem mi)
        {
            try
            {
                if (mi.SenderEmailType == "EX")
                {
                    return mi.Sender?.GetExchangeUser()?.PrimarySmtpAddress;
                }
                return mi.SenderEmailAddress;
            } catch { return null; }
        }

        private static string GetSmtpAddress(Outlook.Recipient r)
        {
            try
            {
                var ae = r.AddressEntry;
                if (ae != null)
                {
                    if (ae.Type == "EX")
                    {
                        var exu = ae.GetExchangeUser();
                        if (exu != null && !string.IsNullOrEmpty(exu.PrimarySmtpAddress)) return exu.PrimarySmtpAddress;
                    }
                    const string PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                    try { var v = ae.PropertyAccessor?.GetProperty(PR_SMTP) as string; if (!string.IsNullOrEmpty(v)) return v; } catch { }
                    var addr = ae.Address; if (!string.IsNullOrEmpty(addr)) return addr;
                }
            } catch { }
            return r.Address;
        }

        private static EntityReference ResolvePartyByEmail(IOrganizationService org, string email)
        {
            if (org == null || string.IsNullOrEmpty(email)) return null;
            var su = FindBySingleEmail(org, "systemuser", "internalemailaddress", email);
            if (su != null) return su;
            var c = FindByAnyEmail(org, "contact", new[] { "emailaddress1", "emailaddress2", "emailaddress3" }, email);
            if (c != null) return c;
            var a = FindByAnyEmail(org, "account", new[] { "emailaddress1", "emailaddress2", "emailaddress3" }, email);
            if (a != null) return a;
            var l = FindByAnyEmail(org, "lead", new[] { "emailaddress1", "emailaddress2", "emailaddress3" }, email);
            if (l != null) return l;
            return null;
        }

        private static EntityReference FindBySingleEmail(IOrganizationService org, string entity, string field, string email)
        {
            try
            {
                var q = new QueryExpression(entity) { ColumnSet = new ColumnSet(false), TopCount = 1 };
                q.Criteria.AddCondition(field, ConditionOperator.Equal, email);
                var r = org.RetrieveMultiple(q);
                if (r != null && r.Entities != null && r.Entities.Count > 0)
                    return r.Entities[0].ToEntityReference();
            } catch { }
            return null;
        }

        private static EntityReference FindByAnyEmail(IOrganizationService org, string entity, string[] fields, string email)
        {
            try
            {
                var q = new QueryExpression(entity) { ColumnSet = new ColumnSet(false), TopCount = 1 };
                var f = new FilterExpression(LogicalOperator.Or);
                foreach (var field in fields) f.AddCondition(field, ConditionOperator.Equal, email);
                q.Criteria.AddFilter(f);
                var r = org.RetrieveMultiple(q);
                if (r != null && r.Entities != null && r.Entities.Count > 0)
                    return r.Entities[0].ToEntityReference();
            } catch { }
            return null;
        }

        
        private static void TryBackfillApptPartiesIfEmpty(IOrganizationService org, Entity existing, Entity update, Outlook.AppointmentItem appt)
        {
            try
            {
                bool hasReq = existing.Contains("requiredattendees") && existing.GetAttributeValue<EntityCollection>("requiredattendees")?.Entities?.Count > 0;
                bool hasOpt = existing.Contains("optionalattendees") && existing.GetAttributeValue<EntityCollection>("optionalattendees")?.Entities?.Count > 0;
                if (!hasReq) update["requiredattendees"] = new EntityCollection(BuildPartiesFromAppointment(appt, true, org));
                if (!hasOpt) update["optionalattendees"] = new EntityCollection(BuildPartiesFromAppointment(appt, false, org));
            } catch { }
        }

        
        /// <summary>
        /// Retourne un nom lisible pour un EntityReference (fullname / name). Fallback: GUID.
        /// </summary>
        private static string GetReadableName(IOrganizationService org, EntityReference er)
        {
            if (er == null) return null;
            if (!string.IsNullOrWhiteSpace(er.Name)) return er.Name;
            try
            {
                string logical = er.LogicalName ?? "";
                string attr = null;
                switch (logical)
                {
                    case "contact": attr = "fullname"; break;
                    case "lead":    attr = "fullname"; break;
                    case "account": attr = "name";     break;
                    case "systemuser": attr = "fullname"; break;
                }
                if (attr != null)
                {
                    var e = org.Retrieve(logical, er.Id, new ColumnSet(attr));
                    var name = e.GetAttributeValue<string>(attr);
                    if (!string.IsNullOrWhiteSpace(name)) return name;
                }
            } catch { }
            return er.Id.ToString();
        }

private static Guid GetOrgId(IOrganizationService org)
        {
            // Remove dependency on Microsoft.Crm.Sdk.Messages.WhoAmI* (not available in this project)
            // and read the OrganizationId directly.
            try
            {
                var q = new QueryExpression("organization")
                {
                    ColumnSet = new ColumnSet("organizationid"),
                    TopCount = 1
                };
                var r = org.RetrieveMultiple(q);
                var e = (r != null && r.Entities != null && r.Entities.Count > 0) ? r.Entities[0] : null;
                return (e != null && e.Contains("organizationid")) ? e.GetAttributeValue<Guid>("organizationid") : Guid.Empty;
            }
            catch
            {
                return Guid.Empty;
            }
        }

        private static T Safe<T>(Func<T> getter)
        {
            try { return getter(); } catch { return default(T); }
        }

        
        private static List<Entity> BuildPartiesFromMail(Outlook.MailItem mi, Outlook.OlMailRecipientType type, IOrganizationService org)
        {
            var list = new List<Entity>();
            try
            {
                if (type == Outlook.OlMailRecipientType.olOriginator)
                {
                    // FROM: use Sender SMTP, resolve to CRM
                    string addr = GetSenderSmtp(mi);
                    if (!string.IsNullOrEmpty(addr))
                    {
                        var p = new Entity("activityparty");
                        p["participationtypemask"] = new OptionSetValue(1); // FROM
                        p["addressused"] = addr;
                        var er = ResolvePartyByEmail(org, addr);
                        if (er != null) p["partyid"] = er;
                        list.Add(p);
                    }
                }
                else
                {
                    // TO / CC / BCC
                    int mask = (type == Outlook.OlMailRecipientType.olTo) ? 2 :
                               (type == Outlook.OlMailRecipientType.olCC) ? 3 : 4;
                    Outlook.Recipients recips = mi.Recipients;
                    foreach (Outlook.Recipient r in recips)
                    {
                        try
                        {
                            if (r.Type != (int)type) continue;
                            string addr = GetSmtpAddress(r);
                            if (string.IsNullOrEmpty(addr)) addr = r.Address;
                            if (string.IsNullOrEmpty(addr)) continue;

                            var party = new Entity("activityparty");
                            party["participationtypemask"] = new OptionSetValue(mask);
                            party["addressused"] = addr;
                            var er = ResolvePartyByEmail(org, addr);
                            if (er != null) party["partyid"] = er;
                            list.Add(party);
                        } catch { }
                    }
                }
            } catch { }
            return list;
        }

        private static List<Entity> BuildPartiesFromAppointment(Outlook.AppointmentItem appt, bool required, IOrganizationService org)
        {
            var list = new List<Entity>();
            try
            {
                Outlook.Recipients recips = appt.Recipients;
                foreach (Outlook.Recipient r in recips)
                {
                    try
                    {
                        var needed = required ? (r.Type == (int)Outlook.OlMeetingRecipientType.olRequired) : (r.Type == (int)Outlook.OlMeetingRecipientType.olOptional);
                        if (!needed) continue;
                        string addr = r.AddressEntry != null ? (r.AddressEntry.GetExchangeUser() != null ? r.AddressEntry.GetExchangeUser().PrimarySmtpAddress : r.AddressEntry.Address) : null;
                        if (string.IsNullOrEmpty(addr)) addr = r.Address;
                        var party = new Entity("activityparty");
                        party["participationtypemask"] = new OptionSetValue(needed ? 5 : 6);
                        party["addressused"] = addr;
                        var er = ResolvePartyByEmail(org, addr);
                        if (er != null) party["partyid"] = er;
                        list.Add(party);
                    } catch { }
                }
            } catch { }
            return list;
        }

    }
}