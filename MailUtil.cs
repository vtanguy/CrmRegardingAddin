using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public static class MailUtil
    {
        /// <summary>
        /// Récupère l'InternetMessageId du mail Outlook.
        /// </summary>
        public static string GetInternetMessageId(Outlook.MailItem mi)
        {
            if (mi == null) return null;
            var pa = mi.PropertyAccessor;
            try { return pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") as string; } catch { }
            try { return pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E") as string; } catch { }
            try { return pa.GetProperty("http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/InternetMessageId") as string; } catch { }
            return null;
        }

        /// <summary>
        /// Recherche l'Email CRM par messageid avec un jeu de colonnes minimal.
        /// (Compat avec les appels existants dans CrmActions/RibbonController)
        /// </summary>
        public static Entity FindCrmEmailByMessageId(IOrganizationService org, string internetId)
        {
            if (org == null || string.IsNullOrWhiteSpace(internetId)) return null;

            var q = new QueryExpression("email")
            {
                ColumnSet = new ColumnSet("activityid", "regardingobjectid", "statecode", "statuscode"),
                Criteria = new FilterExpression(LogicalOperator.And),
                NoLock = true,
                TopCount = 2
            };
            q.Criteria.AddCondition("messageid", ConditionOperator.Equal, internetId);

            var res = org.RetrieveMultiple(q);
            return res.Entities.FirstOrDefault();
        }

        /// <summary>
        /// Variante "complète" : retourne aussi subject + parties (from/to/cc/bcc).
        /// Pratique pour alimenter le panneau des liens.
        /// </summary>
        public static Entity FindCrmEmailByMessageIdFull(IOrganizationService org, string internetId)
        {
            if (org == null || string.IsNullOrWhiteSpace(internetId)) return null;

            var q = new QueryExpression("email")
            {
                ColumnSet = new ColumnSet(
                    "activityid",
                    "subject",
                    "regardingobjectid",
                    "statecode",
                    "statuscode",
                    "from",
                    "to",
                    "cc",
                    "bcc"
                ),
                Criteria = new FilterExpression(LogicalOperator.And),
                NoLock = true,
                TopCount = 2
            };
            q.Criteria.AddCondition("messageid", ConditionOperator.Equal, internetId);

            var res = org.RetrieveMultiple(q);
            return res.Entities.FirstOrDefault();
        }
    }
}
