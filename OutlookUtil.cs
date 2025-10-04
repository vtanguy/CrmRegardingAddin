using System;
using System.Configuration;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public static class OutlookUtil
    {
        // Pose la catégorie spéciale si dispo et si activée par config
        public static void TryMarkTrackedCategory(Outlook.MailItem mi)
        {
            try
            {
                var flag = ConfigurationManager.AppSettings["EnableTrackedCategory"];
                if (!string.Equals(flag, "true", StringComparison.OrdinalIgnoreCase))
                    return;

                if (mi == null) return;
                var app = mi.Application;
                if (app == null) return;

                var session = app.Session;
                if (session == null) return;

                // Nom standard : "Tracked to Dynamics 365"
                var catName = "Tracked to Dynamics 365";

                // Existe dans cette boîte ?
                Outlook.Categories cats = session.Categories;
                Outlook.Category cat = null;
                try { cat = cats[catName]; } catch { /* not found */ }
                if (cat == null)
                {
                    // On n'essaie PAS de la créer : côté serveur la catégorie a un ID
                    // spécifique. On sort si absente.
                    return;
                }

                // Ajoute la catégorie si pas déjà présente
                var current = mi.Categories ?? "";
                if (current.IndexOf(catName, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    mi.Categories = string.IsNullOrEmpty(current) ? catName : (current + "," + catName);
                    mi.Save();
                }
            }
            catch { /* silencieux */ }
        }
    }
}
