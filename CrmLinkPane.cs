using System;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public partial class CrmLinkPane : UserControl
    {
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";

        // UDF DASL (string + id)
        private const string DASL_LinkState_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmlinkstate";
        private const string DASL_LinkState_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80BD";

        private const string DASL_CrmId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmid";
        private const string DASL_CrmId_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C4";

        private const string DASL_RegId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingobjectid";
        private const string DASL_RegId_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C9";

        private const string DASL_RegType_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmregardingobjecttypecode";
        private const string DASL_RegType_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80CA";

        private const string DASL_OrgId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmorgid";
        private const string DASL_OrgId_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C5";
        private const string DASL_EntryId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmEntryID";
        private const string DASL_EntryId_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C3";
        private const string DASL_ObjType_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmObjectTypeCode";
        private const string DASL_ObjType_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80D1";
        private const string DASL_RegardingLabel = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/Regarding";

        // Extra string-named aliases (observed on SSS/MS add-in items)
        private const string DASL_LinkState_String2 = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmLinkState";
        private const string DASL_CrmId_String2 = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmId";
        private const string DASL_OrgId_String2 = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmOrgId";

        private const string DASL_RegId_String_Camel = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectId";
        private const string DASL_RegId_String_Old = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingId";

        private const string DASL_RegType_String_Old = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmRegardingObjectType";

        private IOrganizationService _org;
        private Outlook.MailItem _mail;
        private Outlook.AppointmentItem _appt;

        public Action<string, Guid> OnOpenCrm; // callback pour double-clic

        public CrmLinkPane() { InitializeComponent(); } // Designer gère désormais le bouton "Finaliser"

        public void Initialize(IOrganizationService org) { _org = org; }
        public void SetMailItem(Outlook.MailItem mail) { _appt = null; _mail = mail; RefreshLink(); }
        public void SetAppointmentItem(Outlook.AppointmentItem appt) { _mail = null; _appt = appt; RefreshLink(); }

        private void btnRefresh_Click(object sender, EventArgs e) { RefreshLink(); }
        private void lvLinks_DoubleClick(object sender, EventArgs e)
        {
            if (lvLinks.SelectedItems.Count != 1) return;
            var er = lvLinks.SelectedItems[0].Tag as EntityReference;
            if (er == null) return;
            var cb = OnOpenCrm; if (cb != null) cb(er.LogicalName, er.Id);
        }
        private void lvLinks_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (_org == null) { btnUnlink.Enabled = false; btnFinalize.Enabled = false; return; }
                var any = lvLinks.SelectedItems.Count > 0 || lvLinks.Items.Count > 0;
                Outlook.PropertyAccessor pa = null;
                if (_mail != null) pa = PA(_mail);
                else if (_appt != null) pa = PA(_appt);

                double? ls = pa != null ? (GetDoubleAny(pa, DASL_LinkState_String, DASL_LinkState_String2, DASL_LinkState_Id, DASL_LinkState_Id + "0005", DASL_LinkState_Id + "0003")) : null;
                string crmid = pa != null ? ReadString(pa, DASL_CrmId_String, DASL_CrmId_String2, DASL_CrmId_Id) : null;

                bool prepared = ls.HasValue && ls.Value >= 1.0;
                bool finalized = !string.IsNullOrWhiteSpace(crmid) || (ls.HasValue && ls.Value >= 2.0);

                bool isCompose = (_mail != null && !_mail.Sent);
                btnUnlink.Enabled = any && prepared;
                btnFinalize.Enabled = any && prepared && !finalized && !isCompose; // finaliser seulement si pas déjà finalisé
            }
            catch { /* keep UX responsive */ }
        }

        private static Outlook.PropertyAccessor PA(object item)
        {
            try
            {
                if (item is Outlook.MailItem) return ((Outlook.MailItem)item).PropertyAccessor;
                if (item is Outlook.AppointmentItem) return ((Outlook.AppointmentItem)item).PropertyAccessor;
            }
            catch { }
            return null;
        }

        private static double? GetDouble(Outlook.PropertyAccessor pa, string dasl)
        {
            try
            {
                var o = pa.GetProperty(dasl);
                if (o == null) return null;
                if (o is double) return (double)o;
                if (o is int) return (int)o;
                double d; if (double.TryParse(Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture), out d)) return d;
            }
            catch { }
            return null;
        }
        private static double? GetDoubleAny(Outlook.PropertyAccessor pa, params string[] dasls)
        {
            foreach (var d in dasls)
            {
                var v = GetDouble(pa, d);
                if (v.HasValue) return v;
            }
            return null;
        }
        private static string GetString(Outlook.PropertyAccessor pa, string dasl)
        {
            try { var o = pa.GetProperty(dasl); return o == null ? null : Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture); } catch { return null; }
        }
        private static string ReadString(Outlook.PropertyAccessor pa, params string[] dasls)
        {
            foreach (var d in dasls) { var s = GetString(pa, d); if (!string.IsNullOrEmpty(s)) return s; }
            return null;
        }
        private static int? ReadInt(Outlook.PropertyAccessor pa, params string[] dasls)
        {
            foreach (var d in dasls)
            {
                try
                {
                    var o = pa.GetProperty(d); if (o == null) continue;
                    if (o is int) return (int)o;
                    int v; if (int.TryParse(Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture), out v)) return v;
                }
                catch { }
            }
            return null;
        }

        private void RefreshLink()
        {
            lvLinks.BeginUpdate();
            try
            {
                lvLinks.Items.Clear();
                btnUnlink.Enabled = false;
                btnFinalize.Enabled = false;
                // if (_org == null) return;

                if (_mail != null) RefreshMail();
                else if (_appt != null) RefreshAppointment();
            }
            finally
            {
                lvLinks.EndUpdate();
            }
        }

        // === MAIL ===
        private void RefreshMail()
        {
            var pa = PA(_mail);
            var ls = pa != null ? (GetDoubleAny(pa, DASL_LinkState_String, DASL_LinkState_String2, DASL_LinkState_Id, DASL_LinkState_Id + "0005", DASL_LinkState_Id + "0003")) : null;

            // gestion état boutons (préparé/finalisé)
            bool prepared = ls.HasValue && ls.Value >= 1.0;
            string crmidStr2 = pa != null ? ReadString(pa, DASL_CrmId_String, DASL_CrmId_String2, DASL_CrmId_Id) : null;
            bool finalized = (!string.IsNullOrWhiteSpace(crmidStr2)) || (ls.HasValue && ls.Value >= 2.0);
            btnUnlink.Enabled = prepared;
            bool __isCompose = (_mail != null && !_mail.Sent);
            btnFinalize.Enabled = prepared && !finalized && !__isCompose;

            if (!ls.HasValue || ls.Value < 1.0)
            {
                AddRow("Aucun email CRM", "(non lié)", "", null);
                return;
            }

            // 1) prioritaire : retrouver par UDF crmid
            Guid emailId;
            Entity email = null;
            var crmidStr = crmidStr2;
            if (!string.IsNullOrWhiteSpace(crmidStr) && Guid.TryParse(crmidStr.Trim('{', '}'), out emailId))
            {
                try { email = _org.Retrieve("email", emailId, new ColumnSet("subject", "regardingobjectid", "from", "to", "cc", "bcc")); } catch { email = null; }
            }

            // 2) fallback : correlation via messageid
            if (email == null)
            {
                try
                {
                    var msgId = MailUtil.GetInternetMessageId(_mail);
                    email = MailUtil.FindCrmEmailByMessageIdFull(_org, msgId);
                }
                catch { }
            }

            if (email != null)
            {
                AddRow("Email (CRM)", email.GetAttributeValue<string>("subject") ?? "(sans objet)", "email",
                       new EntityReference("email", email.Id));

                AddRegardingFromProps(pa);
AddPartyRows(email, "from");
                AddPartyRows(email, "to");
                AddPartyRows(email, "cc");
                AddPartyRows(email, "bcc");
                return;
            }

            // 3) Repli : résumé à partir des UDF (sans bruit)
            var regIdStr = pa != null ? ReadString(pa, DASL_RegId_String, DASL_RegId_String_Camel, DASL_RegId_String_Old, DASL_RegId_Id) : null;
            var regType = pa != null ? ReadInt(pa, DASL_RegType_String, DASL_RegType_Id) : (ReadInt(pa, DASL_ObjType_String, DASL_ObjType_Id) ?? ReadInt(pa, DASL_RegType_String_Old));
            EntityReference regRef = null;
            if (!string.IsNullOrWhiteSpace(regIdStr) && regType.HasValue)
            {
                Guid rid; if (Guid.TryParse(regIdStr.Trim('{', '}'), out rid))
                {
                    var logical = LogicalNameFromObjectTypeCode(regType.Value);
                    if (!string.IsNullOrEmpty(logical)) regRef = new EntityReference(logical, rid);
                }
            }

            AddRow("Email lié", (ls.HasValue && ls.Value >= 2.0) ? "(CRM déconnecté)" : "(en attente de résolution CRM)", "email", null);
AddRegardingFromProps(pa);
}

        // === APPOINTMENT ===
        private void RefreshAppointment()
        {
            var pa = PA(_appt);
            var ls = pa != null ? (GetDoubleAny(pa, DASL_LinkState_String, DASL_LinkState_String2, DASL_LinkState_Id, DASL_LinkState_Id + "0005", DASL_LinkState_Id + "0003")) : null;

            bool prepared = ls.HasValue && ls.Value >= 1.0;
            string crmidStr2 = pa != null ? ReadString(pa, DASL_CrmId_String, DASL_CrmId_String2, DASL_CrmId_Id) : null;
            bool finalized = (!string.IsNullOrWhiteSpace(crmidStr2)) || (ls.HasValue && ls.Value >= 2.0);
            btnUnlink.Enabled = prepared;
            bool __isCompose = (_mail != null && !_mail.Sent);
            btnFinalize.Enabled = prepared && !finalized && !__isCompose;

            if (!ls.HasValue || ls.Value < 1.0)
            {
                AddRow("Aucun rendez-vous CRM", "(non lié)", "", null);
                return;
            }

            // Lecture des identifiants CRM (variants)
            var crmidStr = crmidStr2;
            var orgidStr = pa != null ? ReadString(pa, DASL_OrgId_String, DASL_OrgId_String2, DASL_OrgId_Id) : null;
            var objType = pa != null ? (ReadInt(pa, DASL_ObjType_String, DASL_ObjType_Id) ?? 0) : 0;

            // 1) prioritaire : si crmid présent -> retrieve direct
            Guid apptId;
            Entity apptCrm = null;
            if (!string.IsNullOrWhiteSpace(crmidStr) && Guid.TryParse(crmidStr.Trim('{', '}'), out apptId))
            {
                try { apptCrm = _org.Retrieve("appointment", apptId, new ColumnSet("subject", "regardingobjectid", "organizer", "requiredattendees", "optionalattendees")); } catch { apptCrm = null; }
            }

            // 2) fallback : globalobjectid
            if (apptCrm == null)
            {
                try
                {
                    var goid = _appt.GlobalAppointmentID;
                    var found = CrmActions.FindCrmAppointmentByGlobalObjectId(_org, goid);
                    if (found != null) apptCrm = found;
                }
                catch { }
            }

            // 3) Affichage
            if (apptCrm != null)
            {
                AddRow("Rendez-vous (CRM)", apptCrm.GetAttributeValue<string>("subject") ?? "(sans objet)", "appointment",
                       new EntityReference("appointment", apptCrm.Id));
                AddRegardingFromProps(pa);
AddPartyRows(apptCrm, "organizer");
                AddPartyRows(apptCrm, "requiredattendees");
                AddPartyRows(apptCrm, "optionalattendees");
                return;
            }

            // 4) Cas SSS : crmid + crmorgid mais pas de regarding*
            if (!string.IsNullOrWhiteSpace(crmidStr) && !string.IsNullOrWhiteSpace(orgidStr) && (objType == 4201 || objType == 0))
            {
                Guid id; if (Guid.TryParse(crmidStr.Trim('{', '}'), out id))
                {
                    AddRow("Rendez-vous (CRM)", "(résolu par UDF)", "appointment", new EntityReference("appointment", id));
                    var regLbl = pa != null ? ReadString(pa, DASL_RegardingLabel) : null;
                    if (!string.IsNullOrWhiteSpace(regLbl))
                        AddRow("Regarding", regLbl, "", null);
                    return;
                }
            }

            // 5) Dernier repli
            AddRow("Rendez-vous lié", (ls.HasValue && ls.Value >= 2.0) ? "(CRM déconnecté)" : "(en attente de résolution CRM)", "appointment", null);
}

        private void AddRegardingFromProps(Outlook.PropertyAccessor pa)
        {
            if (pa == null) return;
            string regIdStr = ReadString(pa, DASL_RegId_String, DASL_RegId_String_Camel, DASL_RegId_String_Old, DASL_RegId_Id);
            int? regType = ReadInt(pa, DASL_RegType_String, DASL_RegType_Id) ?? (ReadInt(pa, DASL_ObjType_String, DASL_ObjType_Id) ?? ReadInt(pa, DASL_RegType_String_Old));
            string label = ReadString(pa, DASL_RegardingLabel);

            Guid rid;
            string logical = regType.HasValue ? LogicalNameFromObjectTypeCode(regType.Value) : null;
            EntityReference er = null;
            if (!string.IsNullOrWhiteSpace(regIdStr) && Guid.TryParse(regIdStr.Trim('{','}'), out rid) && !string.IsNullOrEmpty(logical))
            {
                er = new EntityReference(logical, rid);
            }

            string display = !string.IsNullOrWhiteSpace(label)
                ? label
                : (er != null ? (er.Name ?? logical ?? "") : "(aucun)");
            AddRow("Regarding", display, er != null ? er.LogicalName : "", er);
        }

        private void AddPartyRows(Entity activity, string attr)
        {
            if (!activity.Contains(attr)) return;
            var ec = activity[attr] as EntityCollection;
            if (ec == null || ec.Entities == null) return;

            foreach (var ap in ec.Entities)
            {
                var er = ap.GetAttributeValue<EntityReference>("partyid");
                var addr = ap.GetAttributeValue<string>("addressused");
                var display = er != null ? (er.Name ?? er.LogicalName) : (addr ?? "(partie)");
                AddRow(attr.ToUpperInvariant(), display, er != null ? er.LogicalName : "", er);
            }
        }

        private void AddRow(string role, string name, string entity, EntityReference tag)
        {
            var item = new ListViewItem(new[] { role, name ?? "", entity ?? "", tag != null ? tag.Id.ToString() : "" });
            item.Tag = tag;
            lvLinks.Items.Add(item);
        }

        private static string LogicalNameFromObjectTypeCode(int typeCode)
        {
            switch (typeCode)
            {
                case 1: return "account";
                case 2: return "contact";
                case 3: return "opportunity";
                case 4: return "lead";
                case 112: return "incident";
                case 4201: return "appointment";
                case 4202: return "email";
                default: return null;
            }
        }

        private static string ResolveName(IOrganizationService org, EntityReference er)
        {
            if (er == null) return "";
            if (!string.IsNullOrWhiteSpace(er.Name)) return er.Name;

            try
            {
                var e = org.Retrieve(er.LogicalName, er.Id, new ColumnSet(true));
                foreach (var key in new[] { "name", "fullname", "subject", "title" })
                {
                    if (e.Attributes.ContainsKey(key))
                    {
                        var val = e[key] as string;
                        if (!string.IsNullOrWhiteSpace(val)) return val;
                    }
                }
                return er.LogicalName + " " + er.Id.ToString("B").ToUpperInvariant();
            }
            catch { return er.LogicalName + " " + er.Id.ToString("B").ToUpperInvariant(); }
        }

        private void btnUnlink_Click(object sender, EventArgs e)
        {
            if (_org == null) return;

            if (_mail != null)
            {
                var ask = MessageBox.Show(
                    "Supprimer aussi l'email CRM ?\n\nOui = annuler le lien ET supprimer l'email CRM\nNon = annuler le lien uniquement (conserver dans CRM)\nAnnuler = ne rien faire",
                    "Annuler le lien",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (ask == DialogResult.Cancel) return;
                var deleteCrmEmail = (ask == DialogResult.Yes);
                var keepCrmItem = !deleteCrmEmail;

                try
                {
                    MsCrmCleanup.UnstampMail(_mail, keepCrmItem);
                    CrmActions.UnlinkOrDeleteCrmEmail(_org, _mail, deleteCrmEmail);
                    MsCrmCleanup.TriggerSyncIfPossible();
                    RefreshLink();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossible d’annuler le lien : " + ex.Message, "CRM",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (_appt != null)
            {
                var ask = MessageBox.Show(
                    "Supprimer aussi le rendez-vous CRM ?\n\nOui = annuler le lien ET supprimer le rendez-vous CRM\nNon = annuler le lien uniquement (conserver dans CRM)\nAnnuler = ne rien faire",
                    "Annuler le lien",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (ask == DialogResult.Cancel) return;
                var deleteCrmAppt = (ask == DialogResult.Yes);
                var keepCrmItem = !deleteCrmAppt;

                try
                {
                    MsCrmCleanup.UnstampAppointment(_appt, keepCrmItem);
                    CrmActions.UnlinkOrDeleteCrmAppointment(_org, _appt, deleteCrmAppt);
                    MsCrmCleanup.TriggerSyncIfPossible();
                    RefreshLink();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossible d’annuler le lien : " + ex.Message, "CRM",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnFinalize_Click(object sender, EventArgs e)
        {
            try
            {
                if (_org == null)
                {
                    MessageBox.Show("Non connecté au CRM.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (_mail == null)
                {
                    MessageBox.Show("Ouvrez un email pour finaliser le lien.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Vérifier l'état préparé (LinkState >= 1 et Regarding présent)
                if (!LinkingApi.HasPreparedMailLink(_mail))
                {
                    MessageBox.Show("Aucun lien préparé à finaliser pour cet email.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    LinkingApi.FinalizePreparedMailIfPossible(_org, _mail);
                    MessageBox.Show("Lien CRM finalisé.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshLink();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossible de finaliser le lien CRM pour cet email.\r\n" + ex.Message,
                                    "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch { /* ne jamais planter l'UI */ }
        }
    }
}
