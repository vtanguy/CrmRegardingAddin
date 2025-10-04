using System;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Xrm.Sdk;

namespace CrmRegardingAddin
{
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
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
            try { System.Diagnostics.Debug.WriteLine("[CRMADDIN] GetCustomUI for " + ribbonID); } catch { }
            // Unique XML pour tous les rubans Outlook
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='CrmTab' label='CRM'>
        <group id='CrmGrp' label='Dynamics 365'>
          <labelControl id='lblState' getLabel='GetStateLabel' />
          <button id='btnConnect' label='Connect CRM' size='large' onAction='OnConnectClicked' />
          <button id='btnRegarding' label='Set Regarding (Mail)' size='large' onAction='OnSetRegardingClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnOpen' label='Open in CRM' size='large' onAction='OnOpenCrmClicked' getEnabled='GetEnabledWhenConnected'/>
          <button id='btnSetRegardingAppt' label='Set Regarding (RDV)' size='large' onAction='OnSetRegardingAppointmentClicked' getEnabled='GetEnabledWhenConnected'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            try { System.Diagnostics.Debug.WriteLine("[CRMADDIN] Ribbon loaded"); } catch { }
            RefreshState();
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
                string diag;
                var svc = CrmConn.Connect(out diag);
                if (svc == null)
                {
                    MessageBox.Show("Connexion CRM échouée.\n\n" + (diag ?? "(aucun détail)"),
                        "CRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                SetConnectedService(svc);
                MessageBox.Show("Connexion CRM réussie.", "CRM",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur de connexion : " + ex.Message, "CRM",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSetRegardingClicked(Office.IRibbonControl control)
        {
            var mi = GetCurrentMailItem();
            if (mi == null) { MessageBox.Show("Sélectionne d’abord un email.", "CRM"); return; }
            if (_org == null) { MessageBox.Show("Pas connecté à CRM.", "CRM"); return; }

            try
            {
                using (var dlg = new SearchDialog(_org))
                {
                    if (dlg.ShowDialog() != DialogResult.OK || dlg.SelectedReference == null) return;
                    CrmActions.SetRegarding(_org, dlg.SelectedReference, mi);

                    // Afficher le pane immédiatement si possible
                    TryInvokeCreatePane("CreatePaneForMailIfLinked", mi);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Set Regarding (Mail) a échoué : " + ex.Message, "CRM");
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

                var er = crmEmail.Contains("regardingobjectid") ? crmEmail.GetAttributeValue<Microsoft.Xrm.Sdk.EntityReference>("regardingobjectid") : null;
                if (er != null) OpenCrm(er.LogicalName, er.Id);
                else OpenCrm("email", crmEmail.Id);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ouverture CRM impossible : " + ex.Message, "CRM");
            }
        }

        public void OnSetRegardingAppointmentClicked(Office.IRibbonControl control)
        {
            var appt = GetCurrentAppointmentItem();
            if (appt == null) { MessageBox.Show("Sélectionne d’abord un rendez-vous.", "CRM"); return; }
            if (_org == null) { MessageBox.Show("Pas connecté à CRM.", "CRM"); return; }

            try
            {
                using (var dlg = new SearchDialog(_org))
                {
                    if (dlg.ShowDialog() != DialogResult.OK || dlg.SelectedReference == null) return;
                    CrmActions.SetRegarding(_org, dlg.SelectedReference, appt);

                    TryInvokeCreatePane("CreatePaneForAppointmentIfLinked", appt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Set Regarding (RDV) a échoué : " + ex.Message, "CRM");
            }
        }

        private void TryInvokeCreatePane(string methodName, object item)
        {
            try
            {
                var insp = Globals.ThisAddIn.Application.ActiveInspector();
                if (insp == null) return;
                var ti = Globals.ThisAddIn;
                var mi = ti.GetType().GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (mi != null) mi.Invoke(ti, new object[] { insp, item });
            }
            catch { }
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
    }
}