using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using Microsoft.Xrm.Sdk;
using WinFormsTimer = System.Windows.Forms.Timer;

namespace CrmRegardingAddin
{
    public partial class ThisAddIn
    {
        // === LinkState DASL constants (for gating pane display) ===
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";
        private const string DASL_LinkState_String = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmlinkstate";
        private const string DASL_LinkState_String_Camel = "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/crmLinkState";
        private const string DASL_LinkState_Id     = "http://schemas.microsoft.com/mapi/id/"     + PS_PUBLIC_STRINGS + "/0x80C8";

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
                if (o == null) { try { o = pa.GetProperty(DASL_LinkState_Id); } catch { } }
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

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonController();
        }

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger.Info("Outlook add-in startup");
            try
            {
                _inspectors = this.Application.Inspectors;
                _inspectors.NewInspector += Inspectors_NewInspector;
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
    }
}
