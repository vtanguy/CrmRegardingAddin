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

        // --- Helpers UDF rendez-vous (C#6-compatible) ---
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
                object ls = null, rn = null, cid = null;
                try { ls = pa.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmlinkstate"); } catch { }
                try { rn = pa.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmregardingobject"); } catch { }
                try { cid = pa.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/crmid"); } catch { }
                var u = new ApptUdf();
                if (ls is double) u.LinkState = (double)ls;
                u.RegardingName = rn as string;
                u.CrmId = cid as string;
                return u;
            }
            catch { return null; }
        }        // --- fin helpers UDF ---


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
            RefreshLink();
        }

        public void SetAppointmentItem(Outlook.AppointmentItem appt)
        {
            _mail = null;
            _appt = appt;
            RefreshLink();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshLink();
        }

        private void lvLinks_DoubleClick(object sender, EventArgs e)
        {
            if (lvLinks.SelectedItems.Count != 1) return;
            var tag = lvLinks.SelectedItems[0].Tag as EntityReference;
            if (tag == null) return;
            OnOpenCrm?.Invoke(tag.LogicalName, tag.Id);
        }

        private void RefreshLink()
        {
            lvLinks.BeginUpdate();
            try
            {
                lvLinks.Items.Clear();
                btnUnlink.Enabled = false;
                if (_org == null) return;

                // ----- MAIL -----
                if (_mail != null)
                {
                    var msgId = MailUtil.GetInternetMessageId(_mail);
                    var email = MailUtil.FindCrmEmailByMessageIdFull(_org, msgId);
                    if (email == null)
                    {
                        AddRow("Aucun email CRM", "(non trouvé)", "", null);
                        return;
                    }

                    AddRow("Email (CRM)", email.GetAttributeValue<string>("subject") ?? "(sans objet)", "email",
                           new EntityReference("email", email.Id));

                    var reg = email.GetAttributeValue<EntityReference>("regardingobjectid");
                    if (reg != null)
                        AddRow("Regarding", reg.Name ?? reg.LogicalName, reg.LogicalName, reg);

                    AddPartyRows(email, "from");
                    AddPartyRows(email, "to");
                    AddPartyRows(email, "cc");
                    AddPartyRows(email, "bcc");

                    btnUnlink.Enabled = true;
                    return;
                }

                // ----- RDV -----


                // ===== APPOINTMENT PATH =====
                if (_appt != null)
                {
                    var goid = _appt.GlobalAppointmentID ?? "";
                    Logger.Info("[Pane] Appt path, GlobalAppointmentID=" + goid);

                    // 1) Essai CRM par globalobjectid
                    Entity apptCrm = null;
                    if (_org != null && !string.IsNullOrEmpty(goid))
                    {
                        apptCrm = CrmActions.FindCrmAppointmentByGlobalObjectId(_org, goid);
                    }

                    if (apptCrm == null)
                    {
                        // 2) Fallback: UDF côté Outlook
                        var udf = ReadAppointmentUdf(_appt);

                        // Si un crmid est stocké, tenter un Retrieve direct
                        Guid crmApptId;
                        if (_org != null && udf != null && !string.IsNullOrEmpty(udf.CrmId)
                            && Guid.TryParse(udf.CrmId.Trim('{', '}'), out crmApptId))
                        {
                            try
                            {
                                var e = _org.Retrieve("appointment", crmApptId,
                                    new ColumnSet("subject", "regardingobjectid", "organizer", "requiredattendees", "optionalattendees"));
                                AddRow("Rendez-vous (CRM)", e.GetAttributeValue<string>("subject") ?? "(sans objet)", "appointment",
                                       new EntityReference("appointment", e.Id));

                                var reg = e.GetAttributeValue<EntityReference>("regardingobjectid");
                                if (reg != null) AddRow("Regarding", reg.Name ?? reg.LogicalName, reg.LogicalName, reg);

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

                        // 3) Sinon afficher l'état "lié (en attente)" si UDF présent
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

                    // 4) Affichage CRM normal
                    AddRow("Rendez-vous (CRM)", apptCrm.GetAttributeValue<string>("subject") ?? "(sans objet)", "appointment",
                           new EntityReference("appointment", apptCrm.Id));

                    var reg2 = apptCrm.GetAttributeValue<EntityReference>("regardingobjectid");
                    if (reg2 != null) AddRow("Regarding", reg2.Name ?? reg2.LogicalName, reg2.LogicalName, reg2);

                    AddPartyRows(apptCrm, "organizer");
                    AddPartyRows(apptCrm, "requiredattendees");
                    AddPartyRows(apptCrm, "optionalattendees");

                    btnUnlink.Enabled = true;
                    Logger.Info("[Pane] Appt found and displayed.");
                    return;
                }
            }
            finally
            {
                lvLinks.EndUpdate();
            }
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
                var name = er != null ? (er.Name ?? er.LogicalName) : (addr ?? "(partie)");
                var eRef = er != null ? er : null;
                AddRow(attr.ToUpperInvariant(), name, er != null ? er.LogicalName : "", eRef);
            }
        }

        private void AddRow(string role, string name, string entity, EntityReference tag)
        {
            var item = new ListViewItem(new[] { role, name ?? "", entity ?? "", tag != null ? tag.Id.ToString() : "" });
            item.Tag = tag;
            lvLinks.Items.Add(item);
        }

        private void btnUnlink_Click(object sender, EventArgs e)
        {
            if (_org == null) return;

            // MAIL
            if (_mail != null)
            {
                var ask = MessageBox.Show(
                    "Supprimer aussi l'email CRM ?\n\nOui = annuler le lien ET supprimer l'email CRM\nNon = annuler le lien uniquement\nAnnuler = ne rien faire",
                    "Annuler le lien",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (ask == DialogResult.Cancel) return;
                var deleteCrmEmail = (ask == DialogResult.Yes);

                try
                {
                    CrmActions.UnlinkOrDeleteCrmEmail(_org, _mail, deleteCrmEmail);
                    RefreshLink();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossible d’annuler le lien : " + ex.Message, "CRM",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            // RDV
            else if (_appt != null)
            {
                var ask = MessageBox.Show(
                    "Supprimer aussi le rendez-vous CRM ?\n\nOui = annuler le lien ET supprimer le rendez-vous CRM\nNon = annuler le lien uniquement\nAnnuler = ne rien faire",
                    "Annuler le lien",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (ask == DialogResult.Cancel) return;
                var deleteCrmAppt = (ask == DialogResult.Yes);

                try
                {
//                    CrmActions.UnlinkOrDeleteCrmAppointment(_org, _appt, deleteCrmAppt);
                    CrmActions.UnlinkOrDeleteCrmAppointment(_org, _appt, true);
                    RefreshLink();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossible d’annuler le lien : " + ex.Message, "CRM",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void lvLinks_SelectedIndexChanged(object sender, EventArgs e) { }
    }
}
