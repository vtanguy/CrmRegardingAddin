using System;
using System.Globalization;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.Drawing;
using System.Drawing.Imaging;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CrmRegardingAddin
{
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
        // === CRM DASL constants for link detection ===
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";
        private const string DASL_LinkState_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmlinkstate";
        private const string DASL_LinkState_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C8";
        private const string DASL_CrmId_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmid";
        private const string DASL_CrmId_Id = "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/0x80C4";

        private static double? TryGetDoubleFromAccessor(Outlook.PropertyAccessor pa, string path)
        {
            if (pa == null || string.IsNullOrEmpty(path)) return null;
            try
            {
                object o = pa.GetProperty(path);
                if (o == null) return null;
                if (o is double) return (double)o;
                if (o is float) return (double)(float)o;
                if (o is int) return (double)(int)o;
                if (o is short) return (double)(short)o;
                if (o is long) return (double)(long)o;
                if (o is bool) return ((bool)o) ? 1.0 : 0.0;
                var s = o as string;
                double d;
                if (!string.IsNullOrEmpty(s) && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
            }
            catch { }
            return null;
        }

        private static string TryGetStringFromAccessor(Outlook.PropertyAccessor pa, string path)
        {
            if (pa == null || string.IsNullOrEmpty(path)) return null;
            try
            {
                object o = pa.GetProperty(path);
                if (o == null) return null;
                var s = o as string;
                if (!string.IsNullOrWhiteSpace(s)) return s;
                try { s = System.Convert.ToString(o, CultureInfo.InvariantCulture); } catch { }
                return string.IsNullOrWhiteSpace(s) ? null : s;
            }
            catch { return null; }
        }

        private enum ReplaceChoice { Replace = 1, Cancel = 2, OpenOnly = 3 }

        // --- Helpers to safely call internal CrmActions methods if not public ---
        private static EntityReference PromptRegarding(IOrganizationService org)
        {
            try
            {
                var t = typeof(CrmActions);
                var m = t.GetMethod("PromptForRegarding", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                if (m != null)
                {
                    var er = m.Invoke(null, new object[] { org }) as EntityReference;
                    if (er != null) return er;
                }
            }
            catch { /* ignore and fallback */ }

            // Fallback mini-dialog (logic name + guid + name)
            using (var f = new Form())
            {
                f.Text = "Choisir cible (fallback)";
                f.StartPosition = FormStartPosition.CenterParent;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MinimizeBox = false; f.MaximizeBox = false;
                f.Width = 460; f.Height = 220; f.Padding = new Padding(10);

                var tl = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 4, AutoSize = true };
                tl.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                tl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                f.Controls.Add(tl);

                var lbl1 = new System.Windows.Forms.Label { Text = "Logical name:", AutoSize = true, Margin = new Padding(0, 4, 8, 4) };
                var txt1 = new TextBox { Dock = DockStyle.Fill, Text = "account" };
                var lbl2 = new System.Windows.Forms.Label { Text = "GUID:", AutoSize = true, Margin = new Padding(0, 4, 8, 4) };
                var txt2 = new TextBox { Dock = DockStyle.Fill };
                var lbl3 = new System.Windows.Forms.Label { Text = "Nom (optionnel):", AutoSize = true, Margin = new Padding(0, 4, 8, 4) };
                var txt3 = new TextBox { Dock = DockStyle.Fill };

                var pnlBtn = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true };
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true };
                var cancel = new Button { Text = "Annuler", DialogResult = DialogResult.Cancel, AutoSize = true };
                pnlBtn.Controls.Add(ok); pnlBtn.Controls.Add(cancel);
                f.AcceptButton = ok; f.CancelButton = cancel;

                tl.Controls.Add(lbl1, 0, 0); tl.Controls.Add(txt1, 1, 0);
                tl.Controls.Add(lbl2, 0, 1); tl.Controls.Add(txt2, 1, 1);
                tl.Controls.Add(lbl3, 0, 2); tl.Controls.Add(txt3, 1, 2);
                tl.Controls.Add(pnlBtn, 0, 3); tl.SetColumnSpan(pnlBtn, 2);

                if (f.ShowDialog() != DialogResult.OK) return null;
                Guid gid;
                if (!Guid.TryParse(txt2.Text, out gid))
                {
                    MessageBox.Show("GUID invalide.", "CRM"); return null;
                }
                var er = new EntityReference(txt1.Text, gid);
                if (!string.IsNullOrWhiteSpace(txt3.Text)) er.Name = txt3.Text;
                return er;
            }
        }

        private static string GetReadableName(IOrganizationService org, EntityReference er)
        {
            try
            {
                var t = typeof(CrmActions);
                var m = t.GetMethod("GetReadableName", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static, null, new[] { typeof(IOrganizationService), typeof(EntityReference) }, null);
                if (m != null)
                {
                    var val = m.Invoke(null, new object[] { org, er }) as string;
                    if (!string.IsNullOrEmpty(val)) return val;
                }
            }
            catch { /* ignore and fallback */ }
            return !string.IsNullOrEmpty(er.Name) ? er.Name : (er.LogicalName + " {" + er.Id.ToString().ToUpper() + "}");
        }

        // Fenêtre de confirmation soignée
        private ReplaceChoice PromptReplaceLink(bool isMail)
        {
            using (var f = new Form())
            {
                f.Text = "CRM";
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MinimizeBox = false;
                f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.AutoScaleMode = AutoScaleMode.Font;
                f.Font = SystemFonts.MessageBoxFont;
                f.Padding = new Padding(14);
                f.Width = 520;
                f.Height = 190;

                var root = new TableLayoutPanel();
                root.Dock = DockStyle.Fill;
                root.AutoSize = true;
                root.ColumnCount = 2;
                root.RowCount = 2;
                root.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                f.Controls.Add(root);

                var pic = new PictureBox();
                pic.Image = SystemIcons.Question.ToBitmap();
                pic.SizeMode = PictureBoxSizeMode.AutoSize;
                pic.Margin = new Padding(0, 2, 12, 0);
                root.Controls.Add(pic, 0, 0);

                var textPanel = new TableLayoutPanel();
                textPanel.Dock = DockStyle.Fill;
                textPanel.AutoSize = true;
                textPanel.ColumnCount = 1;
                textPanel.RowCount = 2;
                textPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                textPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                textPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                root.Controls.Add(textPanel, 1, 0);

                var title = new System.Windows.Forms.Label();
                title.Text = isMail ? "Mail déjà lié" : "Rendez-vous déjà lié";
                title.Font = new Font(f.Font, FontStyle.Bold);
                title.AutoSize = true;
                title.Margin = new Padding(0, 0, 0, 4);
                textPanel.Controls.Add(title, 0, 0);

                var lbl = new System.Windows.Forms.Label();
                lbl.Text = "Voulez-vous remplacer le lien existant ?";
                lbl.AutoSize = true;
                textPanel.Controls.Add(lbl, 0, 1);

                var buttons = new FlowLayoutPanel();
                buttons.FlowDirection = FlowDirection.RightToLeft;
                buttons.Dock = DockStyle.Fill;
                buttons.AutoSize = true;
                buttons.Padding = new Padding(0, 10, 0, 0);
                root.Controls.Add(buttons, 0, 1);
                root.SetColumnSpan(buttons, 2);

                var btnOpen = new Button();
                btnOpen.Text = "Non et ouvrir";
                btnOpen.DialogResult = DialogResult.Retry;
                btnOpen.AutoSize = true;
                buttons.Controls.Add(btnOpen);

                var btnNo = new Button();
                btnNo.Text = "Non";
                btnNo.DialogResult = DialogResult.No;
                btnNo.AutoSize = true;
                buttons.Controls.Add(btnNo);

                var btnYes = new Button();
                btnYes.Text = "Oui (remplacer)";
                btnYes.DialogResult = DialogResult.OK;
                btnYes.AutoSize = true;
                buttons.Controls.Add(btnYes);

                f.AcceptButton = btnYes;
                f.CancelButton = btnNo;

                var dr = f.ShowDialog();
                if (dr == DialogResult.OK) return ReplaceChoice.Replace;
                if (dr == DialogResult.Retry) return ReplaceChoice.OpenOnly;
                return ReplaceChoice.Cancel;
            }
        }

        private static bool IsItemAlreadyLinked(object item)
        {
            try
            {
                Outlook.PropertyAccessor pa = null;
                var mi = item as Outlook.MailItem;
                var appt = item as Outlook.AppointmentItem;
                if (mi != null) pa = mi.PropertyAccessor;
                else if (appt != null) pa = appt.PropertyAccessor;
                if (pa == null) return false;

                double? st = TryGetDoubleFromAccessor(pa, DASL_LinkState_String);
                if (st == null) st = TryGetDoubleFromAccessor(pa, DASL_LinkState_Id);
                if (st.HasValue && st.Value >= 1.0) return true;

                var crmid = TryGetStringFromAccessor(pa, DASL_CrmId_String);
                if (string.IsNullOrWhiteSpace(crmid)) crmid = TryGetStringFromAccessor(pa, DASL_CrmId_Id);
                if (!string.IsNullOrWhiteSpace(crmid)) return true;
            }
            catch { }
            return false;
        }

        private Office.IRibbonUI _ribbon;
        private IOrganizationService _org;

        public static RibbonController Instance { get; private set; }

        public RibbonController()
        {
            Instance = this;
            try { System.Diagnostics.Debug.WriteLine("[CRMADDIN] RibbonController ctor"); } catch { }
        }

        public string GetCustomUI(string ribbonID)
        {
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabMail'>
        <group id='CrmQuickGroup' label='CRM' insertBeforeMso='GroupFind'>
          <button id='btnCrmLinkHome' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
        </group>
      </tab>
      <tab idMso='TabCalendar'></tab>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnCrmLink' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked'/>
          <button id='btnDisconnect' label='Déconnexion CRM' size='large' onAction='OnCrmDisconnect' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnDiagProps' label='Diagnostiquer propriétés' size='large' onAction='OnDiagPropsClicked'/>
        </group>
      </tab>
    </tabs>
    <contextualTabs>
      <tabSet idMso='TabSetAppointment'>
        <tab idMso='TabAppointment'>
          <group id='CrmQuickGroup_CalCtx' label='CRM'>
            <button id='btnCrmLink_CalCtx' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          </group>
        </tab>
      </tabSet>
    </contextualTabs>
  </ribbon>
</customUI>";
                case "Microsoft.Mapi.Mail.Read":
                case "Microsoft.Outlook.Mail.Read":
                    return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabReadMessage'>
        <group id='CrmQuickGroup_Read' label='CRM'>
          <button id='btnCrmLinkHome_Read' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
        </group>
      </tab>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnCrmLink' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked'/>
          <button id='btnDisconnect' label='Déconnexion CRM' size='large' onAction='OnCrmDisconnect' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnDiagProps' label='Diagnostiquer propriétés' size='large' onAction='OnDiagPropsClicked'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
                case "Microsoft.Outlook.Appointment":
                    return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabAppointment'>
        <group id='CrmQuickGroup_ApptInsp' label='CRM'>
          <button id='btnCrmLink_ApptInsp' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
        </group>
      </tab>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnCrmLink' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked'/>
          <button id='btnDisconnect' label='Déconnexion CRM' size='large' onAction='OnCrmDisconnect' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnDiagProps' label='Diagnostiquer propriétés' size='large' onAction='OnDiagPropsClicked'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
                case "Microsoft.Outlook.MeetingRequest":
                    return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabMeeting'>
        <group id='CrmQuickGroup_MtgInsp' label='CRM'>
          <button id='btnCrmLink_MtgInsp' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
        </group>
      </tab>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnCrmLink' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked'/>
          <button id='btnDisconnect' label='Déconnexion CRM' size='large' onAction='OnCrmDisconnect' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnDiagProps' label='Diagnostiquer propriétés' size='large' onAction='OnDiagPropsClicked'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
                case "Microsoft.Outlook.Mail.Compose":
                case "Microsoft.Mapi.Mail.Compose":   // (pour compatibilité selon versions)
                    return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabNewMailMessage'>
        <!-- Même emplacement logique que les autres : avant un groupe Office standard -->
        <group id='CrmQuickGroup_Compose' label='CRM' insertBeforeMso='GroupInclude'>
          <button id='btnCrmLink_Compose' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
        </group>
      </tab>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnCrmLink' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked'/>
          <button id='btnDisconnect' label='Déconnexion CRM' size='large' onAction='OnCrmDisconnect' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnDiagProps' label='Diagnostiquer propriétés' size='large' onAction='OnDiagPropsClicked'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

              default:
                    return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnCrmLink' label='Lien CRM' size='large' onAction='OnCrmLinkClicked' getImage='GetCrmIcon'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked'/>
          <button id='btnDisconnect' label='Déconnexion CRM' size='large' onAction='OnCrmDisconnect' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnDiagProps' label='Diagnostiquer propriétés' size='large' onAction='OnDiagPropsClicked'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
            }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            try { System.Diagnostics.Debug.WriteLine("[CRMADDIN] Ribbon loaded"); } catch { }
            RefreshState();
        }

        private IOrganizationService EnsureConnectedOrPrompt()
        {
            if (_org != null) return _org;
            try
            {
                using (var dlg = new LoginPromptForm())
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return null;

                    string diag;
                    var svc = CrmConn.ConnectWithCredentials(dlg.EnteredUserName, dlg.EnteredPassword, out diag);
                    if (svc == null)
                    {
                        MessageBox.Show("Connexion CRM échouée.\r\n\r\n" + (diag ?? "(aucun détail)"),
                            "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                    SetConnectedService(svc);
                    return _org;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur de connexion CRM : " + ex.Message, "CRM",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public string GetStateLabel(Office.IRibbonControl control)
        {
            return _org != null ? "CRM: Connecté" : "CRM: Hors ligne";
        }

        public bool GetEnabledWhenConnected(Office.IRibbonControl control)
        {
            return _org != null;
        }

        public void OnConnectClicked(Office.IRibbonControl control)
        {
            try
            {
                var dlg = new LoginPromptForm();
                if (dlg.ShowDialog() != DialogResult.OK) return;

                string diag;
                var svc = CrmConn.ConnectWithCredentials(dlg.EnteredUserName, dlg.EnteredPassword, out diag);
                if (svc == null)
                {
                    MessageBox.Show("Connexion CRM échouée.\r\n\r\n" + (diag ?? "(aucun détail)"),
                        "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                SetConnectedService(svc);
                MessageBox.Show("Connecté à CRM.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur de connexion CRM : " + ex.Message, "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnOpenCrmClicked(Office.IRibbonControl control)
        {
            if (_org == null) { MessageBox.Show("Pas connecté à CRM.", "CRM"); return; }
            var mi = GetCurrentMailItem();
            if (mi == null) { MessageBox.Show("Sélectionne d’abord un email.", "CRM"); return; }

            try
            {
                var msgId = MailUtil.GetInternetMessageId(mi);
                var crmEmail = MailUtil.FindCrmEmailByMessageId(_org, msgId);
                if (crmEmail == null) { MessageBox.Show("Aucun Email CRM trouvé pour ce message.", "CRM"); return; }

                var er = crmEmail.Contains("regardingobjectid") ? crmEmail.GetAttributeValue<EntityReference>("regardingobjectid") : null;
                if (er != null) OpenCrm(er.LogicalName, er.Id);
                else OpenCrm("email", crmEmail.Id);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ouverture CRM impossible : " + ex.Message, "CRM");
            }
        }

        public void OnCrmLinkClicked(Office.IRibbonControl control)
        {
            var org = EnsureConnectedOrPrompt();
            if (org == null) return;

            try
            {
                var app = Globals.ThisAddIn.Application;
                var inspector = app.ActiveInspector();

                if (inspector != null && inspector.CurrentItem != null)
                {
                    var mi = inspector.CurrentItem as Outlook.MailItem;
                    if (mi != null)
                    {
                        if (IsItemAlreadyLinked(mi))
                        {
                            var choice = PromptReplaceLink(true);
                            if (choice == ReplaceChoice.Cancel) return;
                            if (choice == ReplaceChoice.OpenOnly) { try { mi.Display(false); } catch { } return; }
                        }

                        
                        // Fast path: if From/To match a CRM contact (emailaddress1/2/3), offer direct link
                        try
                        {
                            if (TryOfferDirectContactLink(org, mi)) return;
                        }
                        catch { }
    
                        var regarding = PromptRegarding(org);
                        if (regarding == null) return;
                        string readable = GetReadableName(org, regarding);

                        // Prepare (Outlook only)
                        try { LinkingApi.PrepareMailLinkInOutlookStore(org, mi, regarding.Id, regarding.LogicalName, readable); } catch { }
                        try { Globals.ThisAddIn.CreatePaneForMailIfLinked(inspector, mi); } catch { }

                        AutoCommitIfNotCompose(org, mi); // suppression du prompt CRM - Confirmation (mail inspector)
                        return;
                    }

                    var appt = inspector.CurrentItem as Outlook.AppointmentItem;
                    if (appt != null)
                    {
                        if (IsItemAlreadyLinked(appt))
                        {
                            var choice = PromptReplaceLink(false);
                            if (choice == ReplaceChoice.Cancel) return;
                            if (choice == ReplaceChoice.OpenOnly) { try { appt.Display(false); } catch { } return; }
                        }

                        var regarding = PromptRegarding(org);
                        if (regarding == null) return;
                        string readable = GetReadableName(org, regarding);

                        try { LinkingApi.PrepareAppointmentLinkInOutlookStore(org, appt, regarding.Id, regarding.LogicalName, readable); } catch { }
                        try { Globals.ThisAddIn.CreatePaneForAppointmentIfLinked(inspector, appt); } catch { }

                        var dr = MessageBox.Show("— préparation uniquement, aucun enregistrement CRM maintenant —\r\n(Choisir 'Non' gardera seulement la préparation locale.)",
                                                 "CRM - Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dr == DialogResult.Yes)
                        {
                            try
                            {
                                var id = LinkingApi.CommitAppointmentLinkToCrm(org, appt);
                                LinkingApi.FinalizeAppointmentLinkInOutlookStoreAfterCrmCommit(org, appt, id);
                                MessageBox.Show("Lien rendez-vous enregistré dans le CRM.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex) { MessageBox.Show("Échec de l'enregistrement CRM:\r\n" + ex.Message, "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        }
                        return;
                    }
                }

                var explorer = app.ActiveExplorer();
                var sel = explorer != null ? explorer.Selection : null;
                if (sel != null && sel.Count > 0)
                {
                    var miSel = sel[1] as Outlook.MailItem;
                    if (miSel != null)
                    {
                        if (IsItemAlreadyLinked(miSel))
                        {
                            var choice = PromptReplaceLink(true);
                            if (choice == ReplaceChoice.Cancel) return;
                            if (choice == ReplaceChoice.OpenOnly) { try { miSel.Display(false); } catch { } return; }
                        }

                        
                        // Fast path: if From/To match a CRM contact (emailaddress1/2/3), offer direct link
                        try
                        {
                            if (TryOfferDirectContactLink(org, miSel)) return;
                        }
                        catch { }
                        var regarding = PromptRegarding(org);
        
                        if (regarding == null) return;
                        string readable = GetReadableName(org, regarding);

                        try { LinkingApi.PrepareMailLinkInOutlookStore(org, miSel, regarding.Id, regarding.LogicalName, readable); } catch { }

                        Outlook.Inspector insp = null;
                        try { insp = miSel.GetInspector; } catch { }
                        try { if (insp == null) { miSel.Display(false); insp = miSel.GetInspector; } } catch { }
                        try { if (insp != null) Globals.ThisAddIn.CreatePaneForMailIfLinked(insp, miSel); } catch { }

                        AutoCommitIfNotCompose(org, miSel); // suppression du prompt CRM - Confirmation (mail explorer)
                        return;
                    }

                    var apptSel = sel[1] as Outlook.AppointmentItem;
                    if (apptSel != null)
                    {
                        if (IsItemAlreadyLinked(apptSel))
                        {
                            var choice = PromptReplaceLink(false);
                            if (choice == ReplaceChoice.Cancel) return;
                            if (choice == ReplaceChoice.OpenOnly) { try { apptSel.Display(false); } catch { } return; }
                        }

                        var regarding = PromptRegarding(org);
                        if (regarding == null) return;
                        string readable = GetReadableName(org, regarding);

                        try { LinkingApi.PrepareAppointmentLinkInOutlookStore(org, apptSel, regarding.Id, regarding.LogicalName, readable); } catch { }

                        Outlook.Inspector insp = null;
                        try { insp = apptSel.GetInspector; } catch { }
                        try { if (insp == null) { apptSel.Display(false); insp = apptSel.GetInspector; } } catch { }
                        try { if (insp != null) Globals.ThisAddIn.CreatePaneForAppointmentIfLinked(insp, apptSel); } catch { }

                        var dr = MessageBox.Show("— préparation uniquement, aucun enregistrement CRM maintenant —\r\n(Choisir 'Non' gardera seulement la préparation locale.)",
                                                 "CRM - Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dr == DialogResult.Yes)
                        {
                            try
                            {
                                var id = LinkingApi.CommitAppointmentLinkToCrm(org, apptSel);
                                LinkingApi.FinalizeAppointmentLinkInOutlookStoreAfterCrmCommit(org, apptSel, id);
                                MessageBox.Show("Lien rendez-vous enregistré dans le CRM.", "CRM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex) { MessageBox.Show("Échec de l'enregistrement CRM:\r\n" + ex.Message, "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        }
                        return;
                    }
                }

                MessageBox.Show("Sélectionne ou ouvre un mail ou un rendez-vous.", "CRM");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Échec de création du lien CRM : " + ex.Message, "CRM");
            }
        }

        public void OnCrmDisconnect(Office.IRibbonControl control)
        {
            try
            {
                _org = null;
                RefreshState();
                MessageBox.Show("Vous êtes maintenant déconnecté du CRM.", "CRM",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        public void OnDiagPropsClicked(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                string dump = null;

                var insp = app.ActiveInspector();
                if (insp != null && insp.CurrentItem != null)
                {
                    var mi = insp.CurrentItem as Outlook.MailItem;
                    if (mi != null) dump = CrmMapiInterop.DumpCrmProps(mi);
                    var appt = insp.CurrentItem as Outlook.AppointmentItem;
                    if (dump == null && appt != null) dump = CrmMapiInterop.DumpCrmProps(appt);
                }

                if (dump == null)
                {
                    var explorer = app.ActiveExplorer();
                    var sel = explorer != null ? explorer.Selection : null;
                    if (sel != null && sel.Count > 0)
                    {
                        var miSel = sel[1] as Outlook.MailItem;
                        if (miSel != null) dump = CrmMapiInterop.DumpCrmProps(miSel);
                        var apptSel = sel[1] as Outlook.AppointmentItem;
                        if (dump == null && apptSel != null) dump = CrmMapiInterop.DumpCrmProps(apptSel);
                    }
                }

                if (string.IsNullOrEmpty(dump)) dump = "Aucun mail/rdv sélectionné.";
                ShowLargeText("CRM UDF (diagnostic)", dump);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Diagnostic impossible : " + ex.Message, "CRM");
            }
        }

        private void ShowLargeText(string title, string text)
        {
            var f = new Form();
            f.Text = title;
            f.Width = 900;
            f.Height = 700;
            var tb = new TextBox();
            tb.Multiline = true;
            tb.ReadOnly = true;
            tb.ScrollBars = ScrollBars.Both;
            tb.WordWrap = false;
            tb.Dock = DockStyle.Fill;
            tb.Font = new Font("Consolas", 10);
            tb.Text = text;
            f.Controls.Add(tb);
            f.StartPosition = FormStartPosition.CenterParent;
            f.ShowDialog();
        }

        private Outlook.MailItem GetCurrentMailItem()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var explorer = app.ActiveExplorer();
                var sel = explorer.Selection;
                if (sel != null && sel.Count > 0) return sel[1] as Outlook.MailItem;
            }
            catch { }
            return null;
        }

        private Outlook.AppointmentItem GetCurrentAppointmentItem()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var explorer = app.ActiveExplorer();
                var sel = explorer.Selection;
                if (sel != null && sel.Count > 0) return sel[1] as Outlook.AppointmentItem;
            }
            catch { }
            return null;
        }

        // === API attendue par ThisAddIn ===
        public void SetConnectedService(IOrganizationService svc)
        {
            _org = svc;
            try { System.Diagnostics.Debug.WriteLine("[CRMADDIN] SetConnectedService: " + (_org != null)); } catch { }
            RefreshState();
        }

        public void RefreshState()
        {
            try { _ribbon?.Invalidate(); } catch { }
        }

        public void OpenCrm(string logicalName, Guid id)
        {
            try
            {
                var baseUrl = ConfigurationManager.AppSettings["CrmBaseUrl"];
                if (string.IsNullOrEmpty(baseUrl))
                {
                    MessageBox.Show("CrmBaseUrl manquant dans App.config", "CRM");
                    return;
                }
                var url = string.Format("{0}/main.aspx?etn={1}&pagetype=entityrecord&id=%7B{2}%7D",
                    baseUrl.TrimEnd('/'), logicalName, id.ToString().ToUpper());
                System.Diagnostics.Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible d’ouvrir CRM : " + ex.Message, "CRM");
            }
        }

        public IOrganizationService Org { get { return _org; } }

        // --- Icon support (custom image from embedded resource) ---
        private stdole.IPictureDisp _crmIconOnline;
        private stdole.IPictureDisp _crmIconOffline;

        
        private static Bitmap MakeGrayscale(Bitmap original)
        {
            var gray = new Bitmap(original.Width, original.Height);
            using (var g = Graphics.FromImage(gray))
            {
                var cm = new ColorMatrix(new float[][] {
                    new float[] {.3f, .3f, .3f, 0, 0},
                    new float[] {.59f, .59f, .59f, 0, 0},
                    new float[] {.11f, .11f, .11f, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {0, 0, 0, 0, 1}
                });
                using (var ia = new ImageAttributes())
                {
                    ia.SetColorMatrix(cm);
                    g.DrawImage(original, new Rectangle(0,0,original.Width, original.Height), 0,0,original.Width, original.Height, GraphicsUnit.Pixel, ia);
                }
            }
            return gray;
        }

        public stdole.IPictureDisp GetCrmIcon(Office.IRibbonControl control)
        {
            try
            {
                bool connected = (_org != null);
                if (connected && _crmIconOnline != null) return _crmIconOnline;
                if (!connected && _crmIconOffline != null) return _crmIconOffline;

                var asm = Assembly.GetExecutingAssembly();
                using (var s = asm.GetManifestResourceStream("CrmRegardingAddin.Resources.TF_30x32.png"))
                {
                    if (s == null) return null;
                    using (var bmp = new Bitmap(s))
                    {
                        if (connected)
                        {
                            _crmIconOnline = PictureConverter.GetIPictureDisp(new Bitmap(bmp));
                            return _crmIconOnline;
                        }
                        else
                        {
                            using (var gray = MakeGrayscale(bmp))
                            {
                                _crmIconOffline = PictureConverter.GetIPictureDisp(new Bitmap(gray));
                                return _crmIconOffline;
                            }
                        }
                    }
                }
            }
            catch { return null; }
        }

        private class PictureConverter : AxHost
        {
            private PictureConverter() : base("") { }
            public static stdole.IPictureDisp GetIPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
            }
        }

        
        // ===== Helpers: mail compose detection + auto-commit for non-compose =====
        private bool IsComposeMail(Outlook.MailItem mi)
        {
            try { return mi != null && !mi.Sent; } catch { return true; }
        }

        private void AutoCommitIfNotCompose(IOrganizationService org, Outlook.MailItem mi)
        {
            // En composition : on ne fait que la préparation locale (pas de commit)
            if (IsComposeMail(mi)) return;
            try
            {
                var id = LinkingApi.CommitMailLinkToCrm(org, mi);
                LinkingApi.FinalizeMailLinkInOutlookStoreAfterCrmCommit(org, mi, id);
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show("Échec de l'enregistrement CRM:\r\n" + ex.Message, "CRM",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                } catch { }
            }
                    var insp = mi != null ? mi.GetInspector : null;
                    if (insp != null) Globals.ThisAddIn.CreatePaneForMailIfLinked(insp, mi);
        }

        #region CRM_DirectLink_Contact_ByEmail

        /// <summary>Offer direct link to a CRM contact if From or any To match a contact (emailaddress1/2/3).</summary>
        private bool TryOfferDirectContactLink(IOrganizationService org, Outlook.MailItem mi)
        {
            try
            {
                if (org == null || mi == null) return false;

                var emails = new System.Collections.Generic.List<string>();

                // FROM first
                var from = GetSenderSmtp(mi);
                if (!string.IsNullOrWhiteSpace(from)) emails.Add(from);

                // Then TO
                foreach (Outlook.Recipient r in mi.Recipients)
                {
                    if (r == null) continue;
                    if (r.Type != (int)Outlook.OlMailRecipientType.olTo) continue;
                    var s = GetRecipientSmtp(r);
                    if (!string.IsNullOrWhiteSpace(s)) emails.Add(s);
                }

                // Deduplicate by case-insensitive compare and keep order
                var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                var ordered = new System.Collections.Generic.List<string>();
                foreach (var s in emails)
                {
                    if (seen.Add(s)) ordered.Add(s);
                }

                foreach (var smtp in ordered)
                {
                    var contact = FindContactByAnyEmail(org, smtp);
                    if (contact == null) continue;

                    var fullname = (contact.Contains("fullname") ? (contact["fullname"] as string) : null) ?? smtp;

                    var dr = MessageBox.Show(
                        "Un contact CRM correspond à l'adresse :\n" + smtp +
                        "\n\nCréer le lien directement vers :\n" + fullname + " ?",
                        "Lien CRM direct vers un contact",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        // Use existing flow to set regarding and optionally commit
                        var regarding = new EntityReference("contact", contact.Id);
                        try
                        {
                            // Prepare local link props + pane (same UX as normal flow)
                            var readable = fullname;
                            LinkingApi.PrepareMailLinkInOutlookStore(org, mi, regarding.Id, regarding.LogicalName, readable);

                            try
                            {
                                Outlook.Inspector insp = null;
                                try { insp = mi.GetInspector; } catch { }
                                if (insp == null) { try { mi.Display(false); insp = mi.GetInspector; } catch { } }
                                if (insp != null) { try { Globals.ThisAddIn.CreatePaneForMailIfLinked(insp, mi); } catch { } }
                            }
                            catch { }

                            AutoCommitIfNotCompose(org, mi); // suppression du prompt CRM - Confirmation (lien direct contact)
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Impossible de créer le lien direct :\n" + ex.Message, "CRM",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        return TrueSafe();
                    }
                    else
                    {
                        // User declined direct link; continue to normal search flow
                        return FalseSafe();
                    }
                }
            }
            catch { }
            return FalseSafe();
        }

        private bool TrueSafe() { return true; }
        private bool FalseSafe() { return false; }

        private Microsoft.Xrm.Sdk.Entity FindContactByAnyEmail(IOrganizationService org, string email)
        {
            try
            {
                var q = new QueryExpression("contact")
                {
                    ColumnSet = new ColumnSet("fullname", "contactid"),
                    TopCount = 1
                };
                var or = new FilterExpression(LogicalOperator.Or);
                or.AddCondition("emailaddress1", ConditionOperator.Equal, email);
                or.AddCondition("emailaddress2", ConditionOperator.Equal, email);
                or.AddCondition("emailaddress3", ConditionOperator.Equal, email);
                q.Criteria.AddFilter(or);
                var r = org.RetrieveMultiple(q);
                return (r != null && r.Entities != null && r.Entities.Count > 0) ? r.Entities[0] : null;
            }
            catch { return null; }
        }

        private string GetSenderSmtp(Outlook.MailItem mi)
        {
            try
            {
                var ae = mi?.Sender;
                if (ae != null)
                {
                    try
                    {
                        var exu = ae.GetExchangeUser();
                        if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress))
                            return exu.PrimarySmtpAddress;
                    }
                    catch { }
                    try
                    {
                        var dl = ae.GetExchangeDistributionList();
                        if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress))
                            return dl.PrimarySmtpAddress;
                    }
                    catch { }
                    try
                    {
                        var addr = ae.Address;
                        if (!string.IsNullOrWhiteSpace(addr)) return addr;
                    }
                    catch { }
                }
            }
            catch { }
            try
            {
                if (!string.IsNullOrWhiteSpace(mi?.SenderEmailAddress))
                    return mi.SenderEmailAddress;
            }
            catch { }
            return null;
        }

        private string GetRecipientSmtp(Outlook.Recipient r)
        {
            try
            {
                if (r?.AddressEntry != null)
                {
                    var exu = r.AddressEntry.GetExchangeUser();
                    if (exu != null && !string.IsNullOrWhiteSpace(exu.PrimarySmtpAddress))
                        return exu.PrimarySmtpAddress;
                    var dl = r.AddressEntry.GetExchangeDistributionList();
                    if (dl != null && !string.IsNullOrWhiteSpace(dl.PrimarySmtpAddress))
                        return dl.PrimarySmtpAddress;
                }
            }
            catch { }
            try
            {
                if (!string.IsNullOrWhiteSpace(r?.Address))
                    return r.Address;
            }
            catch { }
            return null;
        }

        #endregion
    }
}
