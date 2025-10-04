using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Xrm.Sdk;
using WinFormsTimer = System.Windows.Forms.Timer;


namespace CrmRegardingAddin
{
    public partial class ThisAddIn
    {
        private Dictionary<Outlook.Inspector, CustomTaskPane> _crmPanes = new Dictionary<Outlook.Inspector, CustomTaskPane>();
        private Outlook.Inspectors _inspectors;
        private bool _attemptedStartupConnect = false;
        private IOrganizationService _startupSvcPending;
        private WinFormsTimer _waitRibbonTimer;

        // --- Ruban CRM forcé ---
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonController();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
//            Logger.ConfigureFromAppConfig();
            Logger.Info("Outlook add-in startup");
            try
            {
                _inspectors = this.Application.Inspectors;
                _inspectors.NewInspector += Inspectors_NewInspector;
            }
            catch { }

            // Afficher la fenêtre de connexion au démarrage (avec Retry/Cancel)
            var t = new WinFormsTimer();
            t.Interval = 300; // laisser Outlook afficher sa fenêtre d'abord
            t.Tick += (s, args) =>
            {
                try { (s as WinFormsTimer).Stop(); (s as WinFormsTimer).Dispose(); } catch { }
                TryShowStartupConnectDialogLoop();
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
        }

        /// <summary>
        /// Affiche StartupConnectForm en boucle tant que l'utilisateur choisit "Réessayer".
        /// Valide uniquement si la propriété Service != null et DialogResult == OK.
        /// </summary>
        private void TryShowStartupConnectDialogLoop()
        {
            if (_attemptedStartupConnect) return;
            _attemptedStartupConnect = true;

            IOrganizationService svc = null;

            while (true)
            {
                StartupConnectForm dlg = null;
                try
                {
                    // Instanciation directe + branchement du délégué ConnectAsync attendu par la Form
                    dlg = new StartupConnectForm();
                    dlg.ConnectAsync = async (CancellationToken ct) =>
                    {
                        string diag = null;
                        IOrganizationService s = null;
                        try
                        {
                            // CrmConn.Connect est synchrone : on le lance dans un Task.Run annulable
                            s = await Task.Run<IOrganizationService>(() =>
                            {
                                ct.ThrowIfCancellationRequested();
                                var service = CrmConn.Connect(out diag);
                                ct.ThrowIfCancellationRequested();
                                return service;
                            }, ct);
                        }
                        catch (OperationCanceledException)
                        {
                            return new ConnectResult { Service = null, Diag = "Annulé" };
                        }
                        catch (System.Exception ex)
                        {
                            return new ConnectResult { Service = null, Diag = ex.Message };
                        }

                        return new ConnectResult { Service = s, Diag = diag };
                    };

                    var res = dlg.ShowDialog();

                    if (res == DialogResult.OK && dlg.Service != null)
                    {
                        svc = dlg.Service;
                        break; // succès
                    }
                    if (res == DialogResult.Retry)
                    {
                        // l'utilisateur demande explicitement de réessayer
                        continue;
                    }
                    // Cancel ou autre : on stoppe la boucle
                    break;
                }
                catch
                {
                    // Si la Form plante, on sort de la boucle pour tenter le fallback
                    break;
                }
                finally
                {
                    try { if (dlg != null) dlg.Dispose(); } catch { }
                }
            }

            if (svc == null)
            {
                // Fallback : tentative silencieuse (mêmes creds que le bouton)
                try
                {
                    string diag;
                    svc = CrmConn.Connect(out diag);
                }
                catch { svc = null; }
            }

            // Si le ruban n'est pas prêt, mémoriser et pousser quand il arrive
            if (RibbonController.Instance == null)
            {
                _startupSvcPending = svc;
                if (_waitRibbonTimer == null)
                {
                    _waitRibbonTimer = new WinFormsTimer();
                    _waitRibbonTimer.Interval = 500;
                    _waitRibbonTimer.Tick += WaitRibbonTimer_Tick;
                }
                _waitRibbonTimer.Start();
            }
            else
            {
                RibbonController.Instance.SetConnectedService(svc);
                if (svc != null) RefreshVisibleInspectorPaneIfLinked();
            }
        }

        private void WaitRibbonTimer_Tick(object sender, EventArgs e)
        {
            if (RibbonController.Instance == null) return;
            _waitRibbonTimer.Stop();
            try
            {
                RibbonController.Instance.SetConnectedService(_startupSvcPending);
                if (_startupSvcPending != null) RefreshVisibleInspectorPaneIfLinked();
            }
            catch { }
            finally
            {
                _startupSvcPending = null;
            }
        }

        private void RefreshVisibleInspectorPaneIfLinked()
        {
            try
            {
                var insp = this.Application.ActiveInspector();
                if (insp == null || insp.CurrentItem == null) return;

                var mail = insp.CurrentItem as Outlook.MailItem;
                if (mail != null) { CreatePaneForMailIfLinked(insp, mail); return; }

                var appt = insp.CurrentItem as Outlook.AppointmentItem;
                if (appt != null) { CreatePaneForAppointmentIfLinked(insp, appt); return; }
            }
            catch { }
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            try
            {
                if (Inspector == null || Inspector.CurrentItem == null) return;

                var mail = Inspector.CurrentItem as Outlook.MailItem;
                if (mail != null) { CreatePaneForMailIfLinked(Inspector, mail); return; }

                var appt = Inspector.CurrentItem as Outlook.AppointmentItem;
                if (appt != null) { CreatePaneForAppointmentIfLinked(Inspector, appt); return; }
            }
            catch { }
        }

        // ----- Afficher/rafraîchir le pane pour un MAIL si lié CRM -----
        public void CreatePaneForMailIfLinked(Outlook.Inspector inspector, Outlook.MailItem mail)
        {
            var rc = RibbonController.Instance;
            if (rc == null || rc.Org == null) return;

            var msgId = MailUtil.GetInternetMessageId(mail);
            var crmEmail = MailUtil.FindCrmEmailByMessageId(rc.Org, msgId);
            if (crmEmail == null) return;

            CustomTaskPane existing;
            if (_crmPanes.TryGetValue(inspector, out existing))
            {
                var ctrlExisting = existing.Control as CrmLinkPane;
                if (ctrlExisting != null)
                {
                    ctrlExisting.Initialize(rc.Org);
                    ctrlExisting.SetMailItem(mail);
                    existing.Visible = true;
                    return;
                }
                try { this.CustomTaskPanes.Remove(existing); } catch { }
                _crmPanes.Remove(inspector);
            }

            var ctrl = new CrmLinkPane();
            ctrl.Initialize(rc.Org);
            ctrl.SetMailItem(mail);
            ctrl.OnOpenCrm = (ln, id) => rc.OpenCrm(ln, id);

            var pane = this.CustomTaskPanes.Add(ctrl, "CRM", inspector);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
            pane.Height = 250;
            pane.Visible = true;

            _crmPanes[inspector] = pane;

            var inspEvents = inspector as Outlook.InspectorEvents_10_Event;
            if (inspEvents != null)
            {
                inspEvents.Close += () =>
                {
                    try
                    {
                        CustomTaskPane toRemove;
                        if (_crmPanes.TryGetValue(inspector, out toRemove))
                        {
                            try { this.CustomTaskPanes.Remove(toRemove); } catch { }
                            _crmPanes.Remove(inspector);
                        }
                    }
                    catch { }
                };
            }
        }

        // ----- Afficher/rafraîchir le pane pour un RDV si lié CRM -----
        public void CreatePaneForAppointmentIfLinked(Outlook.Inspector inspector, Outlook.AppointmentItem appt)
        {
            var rc = RibbonController.Instance;
            if (rc == null || rc.Org == null || inspector == null || appt == null) return;

            var goid = appt.GlobalAppointmentID ?? "";
            var crmAppt = CrmActions.FindCrmAppointmentByGlobalObjectId(rc.Org, goid);
            if (crmAppt == null) return; // pas de lien CRM => pas de pane

            Microsoft.Office.Tools.CustomTaskPane existing;
            if (_crmPanes.TryGetValue(inspector, out existing))
            {
                CrmLinkPane ctrlExisting;
                if (TryGetCrmLinkPane(existing, out ctrlExisting))
                {
                    // Pane encore vivant : on le réutilise/rafraîchit
                    ctrlExisting.Initialize(rc.Org);
                    ctrlExisting.SetAppointmentItem(appt);
                    existing.Visible = true;
                    return;
                }

                // Pane mort/stale => on le retire proprement et on l'oublie
                try { this.CustomTaskPanes.Remove(existing); } catch { }
                _crmPanes.Remove(inspector);
            }

            // (Re)création d'un pane neuf
            var ctrl = new CrmLinkPane();
            ctrl.Initialize(rc.Org);
            ctrl.SetAppointmentItem(appt);
            ctrl.OnOpenCrm = (ln, id) => rc.OpenCrm(ln, id);

            var pane = this.CustomTaskPanes.Add(ctrl, "CRM", inspector);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
            pane.Height = 160;
            pane.Visible = true;

            _crmPanes[inspector] = pane;
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
        public void RefreshPaneForAppointment(Outlook.Inspector insp, Outlook.AppointmentItem appt)
        {
            // Chercher un pane existant pour cette fenêtre
            Microsoft.Office.Tools.CustomTaskPane existing = null;

            foreach (var pane in this.CustomTaskPanes)
            {
                if (pane.Control is CrmLinkPane && pane.Window == insp)
                {
                    existing = pane;
                    break;
                }
            }

            if (existing != null)
            {
                this.CustomTaskPanes.Remove(existing);
            }

            // Recréer le pane si un lien CRM existe encore
            CreatePaneForAppointmentIfLinked(insp, appt);
        }
        private bool TryGetCrmLinkPane(Microsoft.Office.Tools.CustomTaskPane pane, out CrmLinkPane ctrl)
        {
            ctrl = null;
            if (pane == null) return false;
            try
            {
                // Peut lever ObjectDisposedException si le pane est déjà détruit
                ctrl = pane.Control as CrmLinkPane;
                return ctrl != null;
            }
            catch (ObjectDisposedException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }
    }
}
