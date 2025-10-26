using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using WinFormsTimer = System.Windows.Forms.Timer;

namespace CrmRegardingAddin
{
    public partial class ThisAddIn
    {
        // === LinkState DASL constants (for gating pane display) ===
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";
        private const string DASL_LinkState_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmlinkstate";
        private const string DASL_LinkState_String_Camel = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmLinkState";
        private const string DASL_LinkState_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80BD";

        // --- CRM ID DASL constants ---
        private const string DASL_RegardingId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingId";

        // --- CRM Regarding Type DASL constants ---
        private const string DASL_RegardingType_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectType";
        private const string DASL_RegardingType_String_Lower = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingobjecttype";
        private const string DASL_RegardingType_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80CA";

        // Optional display name (best-effort; we still prefer to retrieve the name from CRM)
        private const string DASL_RegardingName_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/Regarding";
        private const string DASL_RegardingName_String_Lower = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/regarding";
        private const string DASL_RegardingId_String_Camel = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingid";
        private const string DASL_RegardingId_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C9";

        private const string DASL_CrmId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmId";
        private const string DASL_CrmId_String_Lower = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmid";

        private static bool IsLinkedByLinkState(object item)
        {
            if (item == null) return false;
            Outlook.PropertyAccessor pa = null;
            try
            {
                var mi = item as Outlook.MailItem;
                if (mi != null) pa = mi.PropertyAccessor;
                var ap = item as Outlook.AppointmentItem;
                if (ap != null) pa = ap.PropertyAccessor;
            }
            catch { }
            if (pa == null) return false;

            try
            {
                object o = null;
                // 1) string-named lowercase
                try { o = pa.GetProperty(DASL_LinkState_String); } catch { }
                // 2) id-based
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_Id); } catch { }
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_Id + "0005"); } catch { } }
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_Id + "0003"); } catch { } } }
                // 3) string-named camelCase
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_String_Camel); } catch { } }
                if (o == null) return false;

                double d;
                if (o is double) d = (double)o;
                else if (o is float) d = (float)o;
                else if (o is int) d = (int)o;
                else
                {
                    double tmp = 0.0;
                    var s = o as string;
                    if (s != null)
                    {
                        // Try current UI culture then invariant
                        if (!double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out tmp) &&
                            !double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out tmp))
                            return false;
                        d = tmp;
                    }
                    else
                    {
                        var f = o as IFormattable;
                        if (f != null)
                        {
                            var fs = f.ToString(null, System.Globalization.CultureInfo.InvariantCulture);
                            if (!double.TryParse(fs, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out tmp))
                                return false;
                            d = tmp;
                        }
                        else
                        {
                            try { d = Convert.ToDouble(o, System.Globalization.CultureInfo.InvariantCulture); }
                            catch { return false; }
                        }
                    }
                }
                return d > 0.0;
            }
            catch { return false; }
        }

        private readonly Dictionary<Outlook.Inspector, CustomTaskPane> _crmPanes = new Dictionary<Outlook.Inspector, CustomTaskPane>();
        private Outlook.Inspectors _inspectors;
        private bool _attemptedStartupConnect = false;
        private IOrganizationService _startupSvcPending;
        private WinFormsTimer _waitRibbonTimer;

        // --- Added: reply detection / inline compose hooks ---
        private Outlook.Explorers _explorers;
        private readonly HashSet<Outlook.Explorer> _hookedExplorers = new HashSet<Outlook.Explorer>();
        private DateTime _lastReplyPopupAt = DateTime.MinValue;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonController();
        }

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        // Detect if a MailItem is likely a *reply* (not a brand new item).
        private static bool IsLikelyReply(Outlook.MailItem mail)
        {
            if (mail == null) return false;
            try
            {
                // Heuristic 1: subject prefix like "RE:" or "Ré:"
                var subj = mail.Subject ?? string.Empty;
                if (System.Text.RegularExpressions.Regex.IsMatch(subj, @"^\s*(RE|Re|Ré)\s*[:：]", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    return true;

                // Heuristic 2: 'In-Reply-To' header (PR_IN_REPLY_TO_ID)
                const string PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001F";
                try
                {
                    var v = mail.PropertyAccessor.GetProperty(PR_IN_REPLY_TO_ID) as string;
                    if (!string.IsNullOrEmpty(v)) return true;
                }
                catch { /* property may not exist yet */ }
            }
            catch { }
            return false;
        }

        // --- CRM LinkState helpers (exactly 2) ---
        private static bool TryGetCrmLinkStateEquals2(object item)
        {
            if (item == null) return false;
            Outlook.PropertyAccessor pa = null;
            try
            {
                var mi = item as Outlook.MailItem;
                if (mi != null) pa = mi.PropertyAccessor;
                var ap = item as Outlook.AppointmentItem;
                if (ap != null) pa = ap.PropertyAccessor;
            }
            catch { }
            if (pa == null) return false;

            try
            {
                object o = null;
                // Try string-named then id-based variants
                try { o = pa.GetProperty(DASL_LinkState_String); } catch { }
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_Id + "0005"); } catch { } } // PT_DOUBLE
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_Id + "0003"); } catch { } } // PT_LONG
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_String_Camel); } catch { } }
                if (o == null) return false;

                double d;
                if (o is double) d = (double)o;
                else if (o is float) d = (float)o;
                else if (o is int) d = (int)o;
                else
                {
                    double tmp = 0.0;
                    var s = o as string;
                    if (s != null)
                    {
                        if (!double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out tmp) &&
                            !double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out tmp))
                            return false;
                        d = tmp;
                    }
                    else
                    {
                        try { d = Convert.ToDouble(o, System.Globalization.CultureInfo.InvariantCulture); }
                        catch { return false; }
                    }
                }
                // Strict requirement: crmLinkState == 2 only
                return Math.Abs(d - 2.0) < 0.0001;
            }
            catch { return false; }
        }

        private static string ResolveCrmTextFromOriginal(Outlook.MailItem original)
        {
            try
            {
                return TryGetCrmLinkStateEquals2(original) ? "mail lié au CRM" : "mail non lié au CRM";
            }
            catch { return "mail non lié au CRM"; }
        }

        private bool IsInDraftsFolder(Outlook.MailItem mi)
        {
            try
            {
                if (mi == null) return false;
                var parent = mi.Parent as Outlook.MAPIFolder;
                if (parent == null) return false;
                var drafts = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts) as Outlook.MAPIFolder;
                if (drafts == null) return false;
                return string.Equals(parent.EntryID, drafts.EntryID, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private string BuildComposeKey(Outlook.MailItem mi)
        {
            if (mi == null) return null;
            try
            {
                // Prefer stable EntryID when available (after first save)
                var id = mi.EntryID;
                if (!string.IsNullOrEmpty(id)) return "E:" + id;
            }
            catch { }
            try
            {
                var subj = mi.Subject ?? "";
                var ticks = 0L;
                try { ticks = mi.CreationTime.Ticks; } catch { }
                return "T:" + ticks.ToString() + "|" + subj;
            }
            catch { return null; }
        }

        private static void GetCrmIdsFromOriginal(Outlook.MailItem original, out string crmId, out string crmRegardingId)
        {
            crmId = null; crmRegardingId = null;
            if (original == null) return;
            Outlook.PropertyAccessor pa = null;
            try { pa = original.PropertyAccessor; } catch { }
            if (pa == null) return;

            // crmRegardingId
            try
            {
                object o = null;
                try { o = pa.GetProperty(DASL_RegardingId_String); } catch { }
                if (o == null) { try { o = pa.GetProperty(DASL_RegardingId_String_Camel); } catch { } }
                if (o == null) { try { o = pa.GetProperty(DASL_RegardingId_Id + "001F"); } catch { } } // PT_UNICODE
                if (o == null) { try { o = pa.GetProperty(DASL_RegardingId_Id + "001E"); } catch { } } // PT_STRING8
                crmRegardingId = o as string;
            }
            catch { }

            // crmId (best-effort: common string names)
            try
            {
                object o = null;
                try { o = pa.GetProperty(DASL_CrmId_String); } catch { }
                if (o == null) { try { o = pa.GetProperty(DASL_CrmId_String_Lower); } catch { } }
                crmId = o as string;
            }
            catch { }
        }

        private static EntityReference TryGetRegardingFromOriginalMailItem(Outlook.MailItem original, IOrganizationService org)
        {
            if (original == null) return null;
            Outlook.PropertyAccessor pa = null;
            try { pa = original.PropertyAccessor; } catch { }
            if (pa == null) return null;

            // Read regarding id
            System.Guid rid;
            string ridStr = null;
            try
            {
                object o = null;
                try { o = pa.GetProperty(DASL_RegardingId_String); } catch { }
                if (o == null) { try { o = pa.GetProperty(DASL_RegardingId_String_Camel); } catch { } }
                if (o == null) { try { o = pa.GetProperty(DASL_RegardingId_Id + "001F"); } catch { } }
                if (o == null) { try { o = pa.GetProperty(DASL_RegardingId_Id + "001E"); } catch { } }
                ridStr = o as string;
            }
            catch { }
            if (string.IsNullOrEmpty(ridStr) || !System.Guid.TryParse(ridStr, out rid)) return null;

            // Read type string if present (e.g., 'account' or 'contact')
            string typeStr = null;
            try
            {
                object t = null;
                try { t = pa.GetProperty(DASL_RegardingType_String); } catch { }
                if (t == null) { try { t = pa.GetProperty(DASL_RegardingType_String_Lower); } catch { } }
                if (t == null) { try { t = pa.GetProperty(DASL_RegardingType_Id + "001F"); } catch { } }
                if (t == null) { try { t = pa.GetProperty(DASL_RegardingType_Id + "001E"); } catch { } }
                typeStr = t as string;
            }
            catch { }

            // If we have the type, use it directly
            if (!string.IsNullOrEmpty(typeStr))
            {
                var ln = typeStr.Trim().ToLowerInvariant();
                if (ln == "account" || ln == "contact")
                    return new EntityReference(ln, rid);
            }

            // Otherwise, if org available, try validating contact then account
            if (org != null)
            {
                try { var c = org.Retrieve("contact", rid, new ColumnSet(false)); if (c != null) return new EntityReference("contact", rid); } catch { }
                try { var a = org.Retrieve("account", rid, new ColumnSet(false)); if (a != null) return new EntityReference("account", rid); } catch { }
            }

            // If all else fails, return null; (we avoid guessing a wrong logical name)
            return null;
        }

        private static string GetPrimaryNameForEntity(IOrganizationService org, EntityReference er)
        {
            if (org == null || er == null) return null;
            try
            {
                string[] nameAttrs;
                switch ((er.LogicalName ?? "").ToLowerInvariant())
                {
                    case "contact": nameAttrs = new[] { "fullname" }; break;
                    case "account": nameAttrs = new[] { "name" }; break;
                    default: nameAttrs = new[] { "name", "fullname", "title" }; break;
                }
                var ent = org.Retrieve(er.LogicalName, er.Id, new ColumnSet(nameAttrs));
                if (ent == null) return null;
                foreach (var attr in nameAttrs)
                {
                    if (ent.Attributes.ContainsKey(attr) && ent[attr] is string)
                        return (string)ent[attr];
                }
            }
            catch { }
            return null;
        }

        private void ShowReplyPopupSoon(string crmText)
        {
            if ((DateTime.UtcNow - _lastReplyPopupAt).TotalMilliseconds < 500) return;
            _lastReplyPopupAt = DateTime.UtcNow;

            var timer = new WinFormsTimer();
            timer.Interval = 150;
            timer.Tick += (s, e) =>
            {
                try { var t = (WinFormsTimer)s; t.Stop(); t.Dispose(); } catch { }
                try { MessageBox.Show("reponse créée — " + crmText, "CrmRegardingAddin", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
            };
            try { timer.Start(); } catch { }
        }

        private void AttachExplorer(Outlook.Explorer expl)
        {
            if (expl == null) return;
            if (_hookedExplorers.Contains(expl)) return;
            try
            {
                expl.InlineResponse += Explorer_InlineResponse;
                _hookedExplorers.Add(expl);
            }
            catch { }
        }

        private void DetachExplorer(Outlook.Explorer expl)
        {
            if (expl == null) return;
            if (!_hookedExplorers.Contains(expl)) return;
            try
            {
                expl.InlineResponse -= Explorer_InlineResponse;
            }
            catch { }
            try { _hookedExplorers.Remove(expl); } catch { }
        }

        private void Explorers_NewExplorer(Outlook.Explorer Explorer)
        {
            try { AttachExplorer(Explorer); } catch { }
        }

        private void Explorer_InlineResponse(object Item)
        {
            try
            {
                var mail = Item as Outlook.MailItem;
                if (mail != null && IsLikelyReply(mail))
                {
                    Outlook.MailItem original = null;
                    try
                    {
                        var ex = Application.ActiveExplorer();
                        if (ex != null && ex.Selection != null && ex.Selection.Count > 0)
                            original = ex.Selection[1] as Outlook.MailItem;
                    }
                    catch { }
                    // Skip if reply (compose) is already in Drafts
                    if (IsInDraftsFolder(mail)) return;

                    // Variables pour PREPARE-AND-PANE
                    EntityReference _erChosen = null;
                    string _erChosenName = null;

                    var crmText = ResolveCrmTextFromOriginal(original);
                    var crmIdVal = (string)null; var crmRegardingVal = (string)null;
                    try
                    {
                        GetCrmIdsFromOriginal(original, out crmIdVal, out crmRegardingVal);

                        // 1) Try regarding from original message properties
                        try
                        {
                            IOrganizationService org = null;
                            try { org = RibbonController.Instance != null ? RibbonController.Instance.Org : null; } catch { }
                            var erLink = TryGetRegardingFromOriginalMailItem(original, org);
                            if (erLink != null)
                            {
                                var name = GetPrimaryNameForEntity(org, erLink);
                                if (string.IsNullOrEmpty(name))
                                {
                                    // best-effort from local stored name if available
                                    try
                                    {
                                        var pa = original.PropertyAccessor;
                                        object n = null;
                                        try { n = pa.GetProperty(DASL_RegardingName_String); } catch { }
                                        if (n == null) { try { n = pa.GetProperty(DASL_RegardingName_String_Lower); } catch { } }
                                        name = n as string;
                                    }
                                    catch { }
                                }
                                crmText += " — lien proposé: " + erLink.Id.ToString() + (string.IsNullOrEmpty(name) ? "" : " (" + name + ")");
                                _erChosen = erLink; _erChosenName = name;
                            }
                            else
                            {
                                // 2) Fallback by crmid -> find account/contact via parties
                                System.Guid actId;
                                if (!string.IsNullOrEmpty(crmIdVal) && System.Guid.TryParse(crmIdVal, out actId) && org != null)
                                {
                                    var erWho = CrmActions.GetAccountOrContactFromCrmId(org, actId);
                                    if (erWho != null)
                                    {
                                        var disp = GetPrimaryNameForEntity(org, erWho);
                                        crmText += " — lien proposé: " + erWho.Id.ToString() + (string.IsNullOrEmpty(disp) ? "" : " (" + disp + ")");
                                        _erChosen = erWho; _erChosenName = disp;
                                    }
                                }
                            }
                        }
                        catch { }

                        // (conservation des infos techniques si demandées plus tôt)
                        if (!string.IsNullOrEmpty(crmIdVal)) crmText += " — crmid: " + crmIdVal;
                        if (!string.IsNullOrEmpty(crmRegardingVal)) crmText += " — crmregardingid: " + crmRegardingVal;
                    }
                    catch { }

                    ShowReplyPopupSoon(crmText);

                    // === PREPARE-AND-PANE after popup (inline compose) ===
                    if (_erChosen != null)
                    {
                        var postTimer = new WinFormsTimer();
                        postTimer.Interval = 350;
                        postTimer.Tick += (s2, e2) =>
                        {
                            try { var t2 = (WinFormsTimer)s2; t2.Stop(); t2.Dispose(); } catch { }
                            try
                            {
                                IOrganizationService org = null;
                                try { org = RibbonController.Instance != null ? RibbonController.Instance.Org : null; } catch { }
                                if (org != null)
                                {
                                    string readable = _erChosenName;
                                    if (string.IsNullOrEmpty(readable))
                                        readable = GetPrimaryNameForEntity(org, _erChosen) ?? _erChosen.Name;

                                    try { LinkingApi.PrepareMailLinkInOutlookStore(org, mail, _erChosen.Id, _erChosen.LogicalName, readable); } catch { }

                                    Outlook.Inspector insp = null;
                                    try { insp = mail.GetInspector; } catch { }
                                    try { if (insp == null) { mail.Display(false); insp = mail.GetInspector; } } catch { }
                                    try { if (insp != null) CreatePaneForMailIfLinked(insp, mail); } catch { }
                                }
                            }
                            catch { }
                        };
                        try { postTimer.Start(); } catch { }
                    }
                }
            }
            catch { }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger.Info("Outlook add-in startup");
            try
            {
                _inspectors = this.Application.Inspectors;
                this.Application.ItemSend += Application_ItemSend;
                _inspectors.NewInspector += Inspectors_NewInspector;

                try
                {
                    _explorers = this.Application.Explorers;
                    if (_explorers != null)
                    {
                        _explorers.NewExplorer += Explorers_NewExplorer;
                        // Hook existing explorers (e.g., main window)
                        foreach (Outlook.Explorer ex in _explorers)
                        {
                            try { AttachExplorer(ex); } catch { }
                        }
                    }
                }
                catch { }

            }
            catch { }

            var t = new WinFormsTimer();
            t.Interval = 300;
            t.Tick += (s, args) =>
            {
                try { (s as WinFormsTimer).Stop(); (s as WinFormsTimer).Dispose(); } catch { }
                TryShowStartupConnectDialogOnce();
            };
            t.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_inspectors != null)
                {
                    _inspectors.NewInspector -= Inspectors_NewInspector;
                    _inspectors = null;
                }
            }
            catch { }
            try { _crmPanes.Clear(); } catch { }
            try
            {
                if (_waitRibbonTimer != null)
                {
                    _waitRibbonTimer.Stop();
                    _waitRibbonTimer.Dispose();
                    _waitRibbonTimer = null;
                }
            }
            catch { }

            try
            {
                if (_explorers != null)
                {
                    _explorers.NewExplorer -= Explorers_NewExplorer;
                    try
                    {
                        foreach (var ex in _hookedExplorers)
                        {
                            try { DetachExplorer(ex); } catch { }
                        }
                    }
                    catch { }
                    _explorers = null;
                }
            }
            catch { }

        }

        private void TryShowStartupConnectDialogOnce()
        {
            if (_attemptedStartupConnect) return;
            _attemptedStartupConnect = true;

            IOrganizationService svc = null;
            try
            {
                var target = CredentialStore.GetDefaultTarget();
                string su, sp;
                if (CredentialStore.TryLoad(target, out su, out sp))
                {
                    string diagSilent;
                    svc = CrmConn.ConnectWithCredentials(su, sp, out diagSilent);
                }
            }
            catch { }

            if (svc == null)
            {
                var dlg = new LoginPromptForm();
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string diag;
                    svc = CrmConn.ConnectWithCredentials(dlg.EnteredUserName, dlg.EnteredPassword, out diag);
                    if (svc != null && dlg.RememberPassword)
                        try { CredentialStore.Save(CredentialStore.GetDefaultTarget(), dlg.EnteredUserName, dlg.EnteredPassword); } catch { }
                    else if (svc != null && !dlg.RememberPassword)
                        try { CredentialStore.Delete(CredentialStore.GetDefaultTarget()); } catch { }

                    if (svc == null && !string.IsNullOrEmpty(diag))
                        MessageBox.Show("Connexion CRM échouée.\r\n\r\n" + diag, "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            try { RibbonController.Instance?.SetConnectedService(svc); } catch { }
        }

        private void WaitRibbonTimer_Tick(object sender, EventArgs e)
        {
            if (RibbonController.Instance != null)
            {
                try
                {
                    _waitRibbonTimer.Stop();
                    RibbonController.Instance.SetConnectedService(_startupSvcPending);
                }
                catch { }
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector insp)
        {
            try
            {
                var item = insp.CurrentItem;
                // Show popup when a reply window is created (compose inspector)
                try
                {
                    var mi = item as Outlook.MailItem;
                    if (mi != null && !mi.Sent && IsLikelyReply(mi))
                    {
                        Outlook.MailItem original = null;
                        try
                        {
                            var ex = Application.ActiveExplorer();
                            if (ex != null && ex.Selection != null && ex.Selection.Count > 0)
                                original = ex.Selection[1] as Outlook.MailItem;
                        }
                        catch { }
                        // Skip if reply (compose) is already in Drafts
                       if (IsInDraftsFolder(mi)) return;

                        // Variables pour PREPARE-AND-PANE
                        EntityReference _erChosen = null;
                        string _erChosenName = null;

                        var crmText = ResolveCrmTextFromOriginal(original);
                        var crmIdVal = (string)null; var crmRegardingVal = (string)null;
                        try
                        {
                            GetCrmIdsFromOriginal(original, out crmIdVal, out crmRegardingVal);

                            // 1) Try regarding from original message properties
                            try
                            {
                                IOrganizationService org = null;
                                try { org = RibbonController.Instance != null ? RibbonController.Instance.Org : null; } catch { }
                                var erLink = TryGetRegardingFromOriginalMailItem(original, org);
                                if (erLink != null)
                                {
                                    var name = GetPrimaryNameForEntity(org, erLink);
                                    if (string.IsNullOrEmpty(name))
                                    {
                                        // best-effort from local stored name if available
                                        try
                                        {
                                            var pa = original.PropertyAccessor;
                                            object n = null;
                                            try { n = pa.GetProperty(DASL_RegardingName_String); } catch { }
                                            if (n == null) { try { n = pa.GetProperty(DASL_RegardingName_String_Lower); } catch { } }
                                            name = n as string;
                                        }
                                        catch { }
                                    }
                                    crmText += " — lien proposé: " + erLink.Id.ToString() + (string.IsNullOrEmpty(name) ? "" : " (" + name + ")");
                                    _erChosen = erLink; _erChosenName = name;
                                }
                                else
                                {
                                    // 2) Fallback by crmid -> find account/contact via parties
                                    System.Guid actId;
                                    if (!string.IsNullOrEmpty(crmIdVal) && System.Guid.TryParse(crmIdVal, out actId) && org != null)
                                    {
                                        var erWho = CrmActions.GetAccountOrContactFromCrmId(org, actId);
                                        if (erWho != null)
                                        {
                                            var disp = GetPrimaryNameForEntity(org, erWho);
                                            crmText += " — lien proposé: " + erWho.Id.ToString() + (string.IsNullOrEmpty(disp) ? "" : " (" + disp + ")");
                                            _erChosen = erWho; _erChosenName = disp;
                                        }
                                    }
                                }
                            }
                            catch { }

                            // (conservation des infos techniques si demandées plus tôt)
                            if (!string.IsNullOrEmpty(crmIdVal)) crmText += " — crmid: " + crmIdVal;
                            if (!string.IsNullOrEmpty(crmRegardingVal)) crmText += " — crmregardingid: " + crmRegardingVal;
                        }
                        catch { }

                        ShowReplyPopupSoon(crmText);

                        // === PREPARE-AND-PANE after popup (reply window) ===
                        if (_erChosen != null)
                        {
                            var postTimer = new WinFormsTimer();
                            postTimer.Interval = 350;
                            postTimer.Tick += (s2, e2) =>
                            {
                                try { var t2 = (WinFormsTimer)s2; t2.Stop(); t2.Dispose(); } catch { }
                                try
                                {
                                    IOrganizationService org = null;
                                    try { org = RibbonController.Instance != null ? RibbonController.Instance.Org : null; } catch { }
                                    if (org != null)
                                    {
                                        string readable = _erChosenName;
                                        if (string.IsNullOrEmpty(readable))
                                            readable = GetPrimaryNameForEntity(org, _erChosen) ?? _erChosen.Name;

                                        try { LinkingApi.PrepareMailLinkInOutlookStore(org, mi, _erChosen.Id, _erChosen.LogicalName, readable); } catch { }
                                        try { CreatePaneForMailIfLinked(insp, mi); } catch { }
                                    }
                                }
                                catch { }
                            };
                            try { postTimer.Start(); } catch { }
                        }
                    }
                }
                catch { }

                if (item is Outlook.MailItem)
                {
                    CreatePaneForMailIfLinked(insp, (Outlook.MailItem)item);
                }
                else if (item is Outlook.AppointmentItem)
                {
                    CreatePaneForAppointmentIfLinked(insp, (Outlook.AppointmentItem)item);
                }
            }
            catch { }
        }

        // === Pane helpers ===

        private CustomTaskPane EnsureCrmPane(Outlook.Inspector insp)
        {
            if (_crmPanes.ContainsKey(insp))
                return _crmPanes[insp];

            object control;
            try
            {
                var paneType = Type.GetType("CrmRegardingAddin.CrmLinkPane");
                if (paneType == null) paneType = FindTypeByName("CrmLinkPane");
                if (paneType == null) return null;

                control = Activator.CreateInstance(paneType);
            }
            catch
            {
                return null;
            }

            CustomTaskPane pane = null;
            try
            {
                var uc = control as System.Windows.Forms.UserControl;
                if (uc == null) return null;

                pane = this.CustomTaskPanes.Add(uc, "CRM", insp);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
                pane.Height = 170;
                pane.Visible = true;

                _crmPanes[insp] = pane;

                try
                {
                    ((Outlook.InspectorEvents_Event)insp).Close += () =>
                    {
                        try
                        {
                            CustomTaskPane toRemove;
                            if (_crmPanes.TryGetValue(insp, out toRemove) && toRemove != null)
                            {
                                try { this.CustomTaskPanes.Remove(toRemove); } catch { }
                                try { _crmPanes.Remove(insp); } catch { }
                            }
                        }
                        catch { }
                    };
                }
                catch { }
            }
            catch { }
            return pane;
        }

        private static Type FindTypeByName(string typeName)
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    var t = asm.GetType("CrmRegardingAddin." + typeName);
                    if (t != null) return t;
                }
                catch { }
            }
            return null;
        }

        private static void SafeInvoke(object target, string method, params object[] args)
        {
            if (target == null) return;
            try
            {
                var mi = target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (mi != null) mi.Invoke(target, args);
            }
            catch { }
        }

        private static void SafeSet(object target, string prop, object value)
        {
            if (target == null) return;
            try
            {
                var pi = target.GetType().GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (pi != null && pi.CanWrite) pi.SetValue(target, value, null);
            }
            catch { }
        }

        private static void SafeSetField(object target, string fieldName, object value)
        {
            if (target == null) return;
            try
            {
                var fi = target.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (fi != null) fi.SetValue(target, value);
            }
            catch { }
        }

        // === Methods used by CrmActions.cs ===

        public void CreatePaneForMailIfLinked(Outlook.Inspector insp, Outlook.MailItem mi)
        {
            try
            {
                // Gate: show pane only if crmlinkstate>0
                if (!IsLinkedByLinkState(mi)) { HidePaneForInspector(insp); return; }
                var pane = EnsureCrmPane(insp);
                if (pane == null) return;

                var ctrl = pane.Control;
                var org1 = RibbonController.Instance != null ? RibbonController.Instance.Org : null;

                SafeInvoke(ctrl, "Initialize", org1);
                SafeInvoke(ctrl, "SetOrganization", org1);
                SafeSet(ctrl, "Organization", org1);

                SafeInvoke(ctrl, "SetMailItem", mi);
                SafeSet(ctrl, "MailItem", mi);

                SafeSetField(ctrl, "OnOpenCrm", new Action<string, Guid>((ln, id) => { try { RibbonController.Instance?.OpenCrm(ln, id); } catch { } }));
                SafeInvoke(ctrl, "RefreshData");
                SafeInvoke(ctrl, "UpdateUI");
            }
            catch { }
        }

        public void CreatePaneForAppointmentIfLinked(Outlook.Inspector insp, Outlook.AppointmentItem appt)
        {
            try
            {
                // Gate: show pane only if crmlinkstate>0
                if (!IsLinkedByLinkState(appt)) { HidePaneForInspector(insp); return; }
                var pane = EnsureCrmPane(insp);
                if (pane == null) return;

                var ctrl = pane.Control;
                var org1 = RibbonController.Instance != null ? RibbonController.Instance.Org : null;

                SafeInvoke(ctrl, "Initialize", org1);
                SafeInvoke(ctrl, "SetOrganization", org1);
                SafeSet(ctrl, "Organization", org1);

                SafeInvoke(ctrl, "SetAppointmentItem", appt);
                SafeSet(ctrl, "AppointmentItem", appt);

                SafeSetField(ctrl, "OnOpenCrm", new Action<string, Guid>((ln, id) => { try { RibbonController.Instance?.OpenCrm(ln, id); } catch { } }));
                SafeInvoke(ctrl, "RefreshData");
                SafeInvoke(ctrl, "UpdateUI");
            }
            catch { }
        }

        private void HidePaneForInspector(Outlook.Inspector insp)
        {
            try
            {
                if (insp == null) return;
                CustomTaskPane pane;
                if (_crmPanes.TryGetValue(insp, out pane) && pane != null)
                {
                    try { this.CustomTaskPanes.Remove(pane); } catch { }
                    try { _crmPanes.Remove(insp); } catch { }
                    return;
                }
                // Fallback: scan CustomTaskPanes for same Window
                try
                {
                    foreach (CustomTaskPane p in this.CustomTaskPanes)
                    {
                        if (object.ReferenceEquals(p.Window, insp))
                        {
                            try { this.CustomTaskPanes.Remove(p); } catch { }
                            break;
                        }
                    }
                }
                catch { }
            }
            catch { }
        }

        // Auto-finalize prepared mail link at send-time if connected; otherwise warn user.
        
        // === Send-time helpers ===
        private static string GetPrimarySmtpAddress(Outlook.Recipient rcpt)
        {
            if (rcpt == null) return null;
            try
            {
                string smtp = null;
                try
                {
                    const string PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                    smtp = rcpt.PropertyAccessor.GetProperty(PR_SMTP) as string;
                } catch { }
                if (!string.IsNullOrEmpty(smtp)) return smtp;

                try
                {
                    var ae = rcpt.AddressEntry;
                    if (ae != null)
                    {
                        if (ae.Type != null && ae.Type.Equals("EX", StringComparison.OrdinalIgnoreCase))
                        {
                            var exu = ae.GetExchangeUser();
                            if (exu != null && !string.IsNullOrEmpty(exu.PrimarySmtpAddress))
                                return exu.PrimarySmtpAddress;
                        }
                        if (!string.IsNullOrEmpty(ae.Address) && ae.Address.Contains("@"))
                            return ae.Address;
                    }
                } catch { }
                return null;
            }
            catch { return null; }
        }

        private static string GetFirstToSmtp(Outlook.MailItem mi)
        {
            if (mi == null) return null;
            try
            {
                Outlook.Recipients recips = mi.Recipients;
                if (recips == null || recips.Count == 0) return null;

                for (int i = 1; i <= recips.Count; i++)
                {
                    Outlook.Recipient r = null;
                    try { r = recips[i]; } catch { }
                    if (r == null) continue;
                    try
                    {
                        if (r.Type == (int)Outlook.OlMailRecipientType.olTo)
                        {
                            var smtp = GetPrimarySmtpAddress(r);
                            if (!string.IsNullOrEmpty(smtp)) return smtp;
                        }
                    } catch { }
                }
                for (int i = 1; i <= recips.Count; i++)
                {
                    Outlook.Recipient r = null;
                    try { r = recips[i]; } catch { }
                    if (r == null) continue;
                    var smtp = GetPrimarySmtpAddress(r);
                    if (!string.IsNullOrEmpty(smtp)) return smtp;
                }
            }
            catch { }
            return null;
        }

        private static bool TryFindCrmContactOrAccountByEmail(IOrganizationService org, string email, out string logicalName, out System.Guid id, out string name)
        {
            logicalName = null; id = System.Guid.Empty; name = null;
            if (org == null || string.IsNullOrEmpty(email)) return false;
            try
            {
                var q1 = new Microsoft.Xrm.Sdk.Query.QueryExpression("contact");
                q1.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("fullname");
                q1.TopCount = 1;
                var f1 = new Microsoft.Xrm.Sdk.Query.FilterExpression(Microsoft.Xrm.Sdk.Query.LogicalOperator.Or);
                f1.AddCondition("emailaddress1", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, email);
                f1.AddCondition("emailaddress2", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, email);
                f1.AddCondition("emailaddress3", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, email);
                q1.Criteria = f1;
                var r1 = org.RetrieveMultiple(q1);
                if (r1 != null && r1.Entities != null && r1.Entities.Count > 0)
                {
                    var e = r1.Entities[0];
                    logicalName = "contact";
                    id = e.Id;
                    if (e.Attributes.Contains("fullname") && e["fullname"] is string) name = (string)e["fullname"];
                    return true;
                }

                var q2 = new Microsoft.Xrm.Sdk.Query.QueryExpression("account");
                q2.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("name");
                q2.TopCount = 1;
                var f2 = new Microsoft.Xrm.Sdk.Query.FilterExpression(Microsoft.Xrm.Sdk.Query.LogicalOperator.Or);
                f2.AddCondition("emailaddress1", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, email);
                f2.AddCondition("emailaddress2", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, email);
                f2.AddCondition("emailaddress3", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, email);
                q2.Criteria = f2;
                var r2 = org.RetrieveMultiple(q2);
                if (r2 != null && r2.Entities != null && r2.Entities.Count > 0)
                {
                    var e = r2.Entities[0];
                    logicalName = "account";
                    id = e.Id;
                    if (e.Attributes.Contains("name") && e["name"] is string) name = (string)e["name"];
                    return true;
                }
            }
            catch { }
            return false;
        }

                // Small inlined dialog to offer "Lier au CRM" / "Envoyer sans lier"
                private static bool ShowLinkChoiceDialog(string displayName, Guid id)
                {
                    try
                    {
                        using (var f = new System.Windows.Forms.Form())
                        {
                            f.Text = "CRM";
                            f.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                            f.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                            f.MinimizeBox = false;
                            f.MaximizeBox = false;
                            f.ShowInTaskbar = false;
                            f.Width = 420;
                            f.Height = 180;

                            var lbl = new System.Windows.Forms.Label();
                            lbl.AutoSize = false;
                            lbl.Left = 12; lbl.Top = 12; lbl.Width = f.ClientSize.Width - 24; lbl.Height = 70;
                            lbl.Text = "Destinataire dans le CRM\r\n" + (string.IsNullOrEmpty(displayName) ? "" : (displayName + "\r\n")) + id.ToString();
                            f.Controls.Add(lbl);

                            var btnLink = new System.Windows.Forms.Button();
                            btnLink.Text = "Lier au CRM";
                            btnLink.Left = f.ClientSize.Width - 2*120 - 20;
                            btnLink.Top = f.ClientSize.Height - 45;
                            btnLink.Width = 120;
                            btnLink.DialogResult = System.Windows.Forms.DialogResult.OK;
                            f.Controls.Add(btnLink);

                            var btnSend = new System.Windows.Forms.Button();
                            btnSend.Text = "Envoyer sans lier";
                            btnSend.Left = f.ClientSize.Width - 120 - 10;
                            btnSend.Top = f.ClientSize.Height - 45;
                            btnSend.Width = 120;
                            btnSend.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                            f.Controls.Add(btnSend);

                            f.AcceptButton = btnLink;
                            f.CancelButton = btnSend;

                            var dr = f.ShowDialog();
                            return dr == System.Windows.Forms.DialogResult.OK;
                        }
                    }
                    catch { return false; }
                }
    private void Application_ItemSend(object item, ref bool cancel)
        {
                                 bool __justPrepared = false;
           try
            {
                var mi = item as Outlook.MailItem;
                if (mi == null) return;

                if (!IsLinkedByLinkState(mi))
                {
                    // If not yet linked/prepared: check first TO recipient against CRM.
                    try
                    {
                        IOrganizationService orgLookup = null;
                        try { orgLookup = RibbonController.Instance != null ? RibbonController.Instance.Org : null; } catch { }
                        var smtp = GetFirstToSmtp(mi);
                        if (!string.IsNullOrEmpty(smtp) && orgLookup != null)
                        {
                            string ln; System.Guid rid; string disp;
                            if (TryFindCrmContactOrAccountByEmail(orgLookup, smtp, out ln, out rid, out disp))
                            {
                                // Offer to link now or send without linking.
                                bool doLink = ShowLinkChoiceDialog(string.IsNullOrEmpty(disp) ? ln : disp, rid);
                                if (doLink)
                                {
                                    try
                                    {
                                        // Prepare now (simulate as-if user linked during compose)
                                        LinkingApi.PrepareMailLinkInOutlookStore(orgLookup, mi, rid, ln, string.IsNullOrEmpty(disp) ? ln : disp);
                                        mi.Save();
                                        __justPrepared = true;
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    catch { }
                }
                bool __hasPrepared = false; try { __hasPrepared = LinkingApi.HasPreparedMailLink(mi); } catch { }

                bool shouldDoCrm = (__justPrepared || __hasPrepared);

                if (shouldDoCrm)
                {
                IOrganizationService org = null;
                /* skip return: we may have just prepared above */
                try { org = RibbonController.Instance != null ? RibbonController.Instance.Org : null; } catch { }
                bool connected = false;
                try
                {
                    if (org != null)
                    {
                        var w = (Microsoft.Crm.Sdk.Messages.WhoAmIResponse)org.Execute(new Microsoft.Crm.Sdk.Messages.WhoAmIRequest());
                        connected = (w != null && w.UserId != Guid.Empty);
                    }
                }
                catch { connected = false; }

                if (!connected)
                {
                    try
                    {
                        MessageBox.Show("Le mail est envoyé mais ne sera pas lié au CRM.\r\nIl faudra finaliser le lien par la suite.",
                                        "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch { }
                    return;
                }

                try
                {
                    // If we just prepared, commit+finalize explicitly; otherwise use existing finalize helper.
                    if (__justPrepared)
                    {
                        try { var __id = LinkingApi.CommitMailLinkToCrm(org, mi); LinkingApi.FinalizeMailLinkInOutlookStoreAfterCrmCommit(org, mi, __id); }
                        catch (Exception ex2)
                        {
                            try { MessageBox.Show("Impossible de lier et d'envoyer (commit/finalize).\r\nLe mail sera envoyé sans lien.\r\n\r\n" + ex2.Message, "CRM", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                        }
                    }
                    else
                    {
                        LinkingApi.FinalizePreparedMailIfPossible(org, mi);
                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        MessageBox.Show("Impossible de finaliser le lien CRM avant l'envoi.\r\n" +
                                        "Le mail sera envoyé sans lien.\r\n" +
                                        "Vous pourrez utiliser 'Finaliser le lien' dans le panneau CRM.\r\n\r\n" +
                                        ex.Message,
                                        "CRM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    catch { }
                }                }

            }
            catch { }
        }
    }
}
