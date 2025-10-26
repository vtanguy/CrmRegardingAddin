using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    /// <summary>
    /// Cleanup of Microsoft Dynamics CRM tracking stamps (compatible with VS2015 PIAs).
    /// Modes:
    ///  - keepCrmItem=true  => "soft unlink": keep CRM data but set crmlinkstate=0, remove tracking categories
    ///  - keepCrmItem=false => "hard unlink": delete CRM named props and categories
    /// </summary>
    internal static class MsCrmCleanup
    {
        private const string PS_PUBLIC_STRINGS = "{00020329-0000-0000-C000-000000000046}";

        private static readonly string[] KnownCrmNames = new[]
        {
            "crmlinkstate",
            "crmid",
            "crmorgid",
            "crmregardingobjectid",
            "crmregardingobjecttypecode",
            "crmregardingobject",
            "crmregardingid",
            "crmregardingobjecttype",
            "crmpartyinfo",
            "crmxml",
            "crmobjecttypecode",
            "crmsecondpagexml",
            "crmasyncsend",
            "crmsss.promotetracker",
            "crmsss_promotetracker",
            "crmssspromotetracker",
            "crmsyncerror",
            "crmhash",
            "crmemailhash",
            "crmconversationhash",
            "crmtrackcategories",
            "mscrm.linkstate",
            "mscrm.regardingobject",
            "mscrm.regardingobjectid",
        };

        private static readonly string[] KnownCrmIds = new[]
        {
            "0x80c3",
            "0x80c4",
            "0x80c5",
            "0x80c8",
            "0x80c9",
            "0x80ca",
            "0x80cf",
            "0x80d1",
            "0x80d4",
            "0x80d5",
            "0x80d7",
            "0x80de",
        };

        private static readonly string[] TrackingCategoryCandidates = new[]
        {
            "Tracked to Dynamics 365",
            "Tracked to Microsoft Dynamics CRM",
            "Tracked to Dynamics CRM",
            "Suivi dans Dynamics 365",
            "Suivi dans Microsoft Dynamics CRM",
            "Suivi vers Dynamics 365",
            "Suivi vers Microsoft Dynamics CRM",
            "Tracked to Dynamics",
            "Suivi dans Dynamics"
        };

        private static string BuildStringDasl(string name)
        {
            return "http://schemas.microsoft.com/mapi/string/" + PS_PUBLIC_STRINGS + "/" + name;
        }

        private static string BuildIdDasl(string hexId)
        {
            return "http://schemas.microsoft.com/mapi/id/" + PS_PUBLIC_STRINGS + "/" + hexId.ToUpperInvariant();
        }

        private static void TryDelete(Outlook.PropertyAccessor pa, string dasl)
        {
            if (pa == null || string.IsNullOrEmpty(dasl)) return;
            try { pa.DeleteProperty(dasl); } catch { }
        }

        private static void TrySetDouble(Outlook.PropertyAccessor pa, string dasl, double value)
        {
            if (pa == null || string.IsNullOrEmpty(dasl)) return;
            try { pa.SetProperty(dasl, value); } catch { }
        }

        private static void RemoveTrackingCategories(object item)
        {
            try
            {
                string cats = null;
                if (item is Outlook.MailItem) cats = ((Outlook.MailItem)item).Categories;
                else if (item is Outlook.AppointmentItem) cats = ((Outlook.AppointmentItem)item).Categories;
                else return;

                if (string.IsNullOrEmpty(cats)) return;

                var tokens = cats.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Select(s => s.Trim())
                                 .ToList();

                bool changed = false;
                for (int i = tokens.Count - 1; i >= 0; i--)
                {
                    var t = tokens[i];
                    if (IsTrackingCategoryName(t))
                    {
                        tokens.RemoveAt(i);
                        changed = true;
                    }
                }
                if (!changed) return;

                var newCats = string.Join("; ", tokens);
                if (item is Outlook.MailItem)
                {
                    ((Outlook.MailItem)item).Categories = newCats;
                }
                else if (item is Outlook.AppointmentItem)
                {
                    ((Outlook.AppointmentItem)item).Categories = newCats;
                }
            }
            catch { }
        }

        private static bool IsTrackingCategoryName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            var n = name.Trim();
            foreach (var c in TrackingCategoryCandidates)
            {
                if (n.Equals(c, StringComparison.OrdinalIgnoreCase)) return true;
            }
            var lo = n.ToLowerInvariant();
            if ((lo.Contains("dynamics") || lo.Contains("crm")) && (lo.Contains("tracked") || lo.Contains("suivi")))
                return true;
            return false;
        }

        private static void DeleteAllCrmProps(Outlook.PropertyAccessor pa)
        {
            if (pa == null) return;

            foreach (var n in KnownCrmNames)
            {
                TryDelete(pa, BuildStringDasl(n));
                TryDelete(pa, BuildStringDasl(n.ToLowerInvariant()));
                TryDelete(pa, BuildStringDasl(n.ToUpperInvariant()));
                TryDelete(pa, BuildStringDasl(n.Replace('.', '_')));
                TryDelete(pa, BuildStringDasl(n.Replace('_', '.')));
            }
            foreach (var id in KnownCrmIds)
            {
                TryDelete(pa, BuildIdDasl(id));
            }
        }

        private static void SetCrmLinkStateZero(Outlook.PropertyAccessor pa)
        {
            TrySetDouble(pa, BuildStringDasl("crmlinkstate"), 0.0);
            TrySetDouble(pa, BuildStringDasl("CRMLINKSTATE"), 0.0);
            TrySetDouble(pa, BuildIdDasl("0x80C8"), 0.0);
        }

        private static void UnstampOneAppointment(Outlook.AppointmentItem appt, bool keepCrmItem, bool save)
        {
            if (appt == null) return;
            try
            {
                var pa = appt.PropertyAccessor;
                if (keepCrmItem)
                {
                    SetCrmLinkStateZero(pa);
                    RemoveTrackingCategories(appt);
                }
                else
                {
                    DeleteAllCrmProps(pa);
                    SetCrmLinkStateZero(pa);
                    RemoveTrackingCategories(appt);
                }
                if (save) { appt.Save(); }
            }
            catch { }
        }

        public static void UnstampAppointment(Outlook.AppointmentItem appt, bool keepCrmItem)
        {
            if (appt == null) return;

            bool wasRecurring = false;
            try { wasRecurring = appt.IsRecurring; } catch { }

            UnstampOneAppointment(appt, keepCrmItem, false);

            if (wasRecurring)
            {
                try
                {
                    var pat = appt.GetRecurrencePattern();
                    if (pat != null)
                    {
                        var master = pat.Parent as Outlook.AppointmentItem;
                        if (master != null && master.EntryID != appt.EntryID)
                        {
                            UnstampOneAppointment(master, keepCrmItem, true);
                        }
                    }
                }
                catch { }
            }

            try { appt.Save(); } catch { }
        }

        private static void UnstampOneMail(Outlook.MailItem mail, bool keepCrmItem, bool save)
        {
            if (mail == null) return;
            try
            {
                var pa = mail.PropertyAccessor;
                if (keepCrmItem)
                {
                    SetCrmLinkStateZero(pa);
                    RemoveTrackingCategories(mail);
                }
                else
                {
                    DeleteAllCrmProps(pa);
                    SetCrmLinkStateZero(pa);
                    RemoveTrackingCategories(mail);
                }
                if (save) mail.Save();
            }
            catch { }
        }

        public static void UnstampMail(Outlook.MailItem mail, bool keepCrmItem)
        {
            UnstampOneMail(mail, keepCrmItem, true);
        }

        public static void TriggerSyncIfPossible()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var ns = app != null ? app.Session : null;
                if (ns != null) { ns.SendAndReceive(false); }
            }
            catch { }
        }
    }
}
