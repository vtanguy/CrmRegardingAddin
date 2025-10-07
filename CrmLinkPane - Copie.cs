
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public partial class CrmLinkPane : UserControl
    {
        private sealed class ApptUdf
        {
            public string CrmId;
            public double? LinkState;
            public string OrgId, RegardingId, RegardingName, RegardingOtc, OwnerSmtp, OwnerUserId;
        }

        private ApptUdf ReadAppointmentUdf(Outlook.AppointmentItem appt)
        {
            try
            {
                if (appt == null) return null;

                var pa = appt.PropertyAccessor;
                Func<string, object> G = n =>
                {
                    try
                    {
                        return pa.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" + n);
                    }
                    catch { return null; }
                };

                var u = new ApptUdf();
                u.LinkState = G("crmlinkstate") as double?;
                u.OrgId = G("crmorgid") as string;
                u.RegardingId = G("crmregardingobjectid") as string;
                u.RegardingName = G("crmregardingobject") as string;
                u.RegardingOtc = G("crmregardingobjecttypecode") as string;
                u.OwnerSmtp = G("crmownersmtp") as string;
                u.OwnerUserId = G("crmownersystemuserid") as string;
                u.CrmId      = G("crmid") as string;
                return u;
            }
            catch { return null; }
        }

        private IOrganizationService _org;
        private Outlook.MailItem _mail;
        private Outlook.AppointmentItem _appt;

        public Action<string, Guid> OnOpenCrm; // callback

        public CrmLinkPane()
        {
            InitializeComponent();
        }

        public void Initialize(IOrganizationService org)
        {
            _org = org;
        }

        public void SetMailItem(Outlook.MailItem mail)
        {
            _appt = null;
            _mail = mail;
            Logger.Info("[Pane] SetMailItem: subject='" + (mail?.Subject ?? "") + "'");
            RefreshLink();
        }

        public void SetAppointmentItem(Outlook.AppointmentItem appt)
        {
            _mail = null;
            _appt = appt;
            Logger.Info("[Pane] SetAppointmentItem: subject='" + (appt?.Subject ?? "") + "'");
            RefreshLink();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            Logger.Info("[Pane] Manual Refresh clicked");
            RefreshLink();
        }

        // Added to match Designer wire-up to avoid CS1061
        private void lvLinks_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnUnlink.Enabled = (lvLinks.SelectedItems.Count > 0);
        }

        private void lvLinks_DoubleClick(object sender, EventArgs e)
        {
            if (lvLinks.SelectedItems.Count != 1) return;
            var tag = lvLinks.SelectedItems[0].Tag as EntityReference;
            if (tag == null) return;
            Logger.Info("[Pane] DoubleClick open CRM: " + tag.LogicalName + " " + tag.Id);
            OnOpenCrm?.Invoke(tag.LogicalName, tag.Id);
        }

        private void AddRow(string type, string text, string logicalName, EntityReference er)
        {
            var it = new ListViewItem(type);
            it.SubItems.Add(text ?? "");
            it.Tag = er;
            lvLinks.Items.Add(it);
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
                var label = (er != null ? (er.Name ?? er.LogicalName) : null) ?? addr ?? "(participant)";
                AddRow(attr, label, er?.LogicalName ?? "", er);
            }
        }

        private void RefreshLink()
        {
            lvLinks.BeginUpdate();
            try
            {
                lvLinks.Items.Clear();
                btnUnlink.Enabled = false;
                if (_org == null) { Logger.Info("[Pane] RefreshLink: _org is null"); return; }

                // ----- MAIL -----
                if (_mail != null)
                {
                    var msgId = MailUtil.GetInternetMessageId(_mail);
                    Logger.Info("[Pane] Mail path, messageId=" + (msgId ?? "(null)"));
                    var email = MailUtil.FindCrmEmailByMessageIdFull(_org, msgId);
                    if (email == null)
                    {
                        Logger.Info("[Pane] Mail not found in CRM");
                        AddRow("Aucun email CRM", "(non trouvé)", "", null);
                        return;
                    }

                    AddRow("Email (CRM)", email.GetAttributeValue<string>("subject") ?? "(sans objet)", "email",
                           new EntityReference("email", email.Id));

                    var reg = email.GetAttributeValue<EntityReference>("regardingobjectid");
                    if (reg != null)
                    {
                        var label = reg.Name ?? reg.LogicalName;
                        AddRow("Regarding", label, reg.LogicalName, reg);
                    }

                    AddPartyRows(email, "from");
                    AddPartyRows(email, "to");
                    AddPartyRows(email, "cc");
                    AddPartyRows(email, "bcc");

                    btnUnlink.Enabled = true;
                    return;
                }

                // ----- RDV -----
                if (_appt != null)
                {
                    var goid = _appt.GlobalAppointmentID ?? "";
                    Logger.Info("[Pane] Appt path, GlobalAppointmentID=" + goid);
                    var appt = CrmActions.FindCrmAppointmentByGlobalObjectId(_org, goid);
                    if (appt == null)
                    {
                        Logger.Info("[Pane] CRM appt not found by GOID. Trying UDF...");
                        var udf = ReadAppointmentUdf(_appt);

                        Guid crmApptId;
                        if (udf != null && !string.IsNullOrEmpty(udf.CrmId) &&
                            Guid.TryParse(udf.CrmId.Trim('{','}'), out crmApptId))
                        {
                            try
                            {
                                var e = _org.Retrieve("appointment", crmApptId,
                                    new ColumnSet("subject","regardingobjectid","organizer","requiredattendees","optionalattendees"));
                                AddRow("Rendez-vous (CRM)", e.GetAttributeValue<string>("subject") ?? "(sans objet)", "appointment",
                                       new EntityReference("appointment", e.Id));

                                var reg = e.GetAttributeValue<EntityReference>("regardingobjectid");
                                if (reg != null)
                                {
                                    var label = reg.Name ?? reg.LogicalName;
                                    AddRow("Regarding", label, reg.LogicalName, reg);
                                }

                                AddPartyRows(e, "organizer");
                                AddPartyRows(e, "requiredattendees");
                                AddPartyRows(e, "optionalattendees");

                                btnUnlink.Enabled = true;
                                Logger.Info("[Pane] CRM appt retrieved by UDF crmid.");
                                return;
                            }
                            catch (Exception ex)
                            {
                                Logger.Info("[Pane] Retrieve by UDF crmid EX: " + ex.Message);
                            }
                        }

                        if (udf != null && (udf.LinkState ?? 0) >= 1.0)
                        {
                            AddRow("Rendez-vous lié (en attente sync)", udf.RegardingName ?? "(regarding)", "", null);
                            Logger.Info("[Pane] Appt linked (pending SSS).");
                        }
                        else
                        {
                            AddRow("Aucun rendez-vous CRM", "(non trouvé)", "", null);
                            Logger.Info("[Pane] Appt not linked.");
                        }
                        return;
                    }

                    AddRow("Rendez-vous (CRM)", appt.GetAttributeValue<string>("subject") ?? "(sans objet)", "appointment",
                           new EntityReference("appointment", appt.Id));

                    var reg2 = appt.GetAttributeValue<EntityReference>("regardingobjectid");
                    if (reg2 != null)
                    {
                        var label2 = reg2.Name ?? reg2.LogicalName;
                        AddRow("Regarding", label2, reg2.LogicalName, reg2);
                    }

                    AddPartyRows(appt, "organizer");
                    AddPartyRows(appt, "requiredattendees");
                    AddPartyRows(appt, "optionalattendees");

                    btnUnlink.Enabled = true;
                    Logger.Info("[Pane] Appt found and displayed.");
                    return;
                }
            }
            finally
            {
                try { lvLinks.EndUpdate(); } catch { }
            }
        }

        private void btnUnlink_Click(object sender, EventArgs e)
        {
            if (_mail != null)
            {
                MessageBox.Show("Délier un mail n'est pas implémenté dans ce panneau.", "CRM");
                return;
            }
            if (_appt != null)
            {
                var r = MessageBox.Show("Souhaitez-vous supprimer le rendez-vous dans le CRM ?",
                    "CRM", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (r == DialogResult.Cancel) return;
                bool deleteCrmAppt = (r == DialogResult.Yes);

                CrmActions.UnlinkOrDeleteCrmAppointment(_org, _appt, deleteCrmAppt);
                Logger.Info("[Pane] btnUnlink clicked. deleteInCrm=" + deleteCrmAppt);
                RefreshLink();
            }
        }
    }
}
