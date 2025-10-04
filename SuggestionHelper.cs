
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CrmRegardingAddin
{
    public static class SuggestionHelper
    {
        public static List<Microsoft.Xrm.Sdk.EntityReference> FindSuggestionsByEmails(IOrganizationService org, IEnumerable<string> emails, int top = 25)
        {
            var results = new List<Microsoft.Xrm.Sdk.EntityReference>();
            var emailList = emails?.Where(e => !string.IsNullOrWhiteSpace(e))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .ToList()
                            ?? new List<string>();
            if (emailList.Count == 0) return results;

            // Contacts: emailaddress1/2/3
            var contactQ = new QueryExpression("contact")
            {
                ColumnSet = new ColumnSet("fullname", "emailaddress1", "emailaddress2", "emailaddress3"),
                TopCount = top
            };
            var contactFilter = new FilterExpression(LogicalOperator.Or);
            foreach (var em in emailList)
            {
                contactFilter.AddCondition("emailaddress1", ConditionOperator.Equal, em);
                contactFilter.AddCondition("emailaddress2", ConditionOperator.Equal, em);
                contactFilter.AddCondition("emailaddress3", ConditionOperator.Equal, em);
            }
            contactQ.Criteria.AddFilter(contactFilter);
            var contacts = org.RetrieveMultiple(contactQ).Entities
                .Select(c =>
                {
                    var er = new Microsoft.Xrm.Sdk.EntityReference("contact", c.Id);
                    er.Name = c.GetAttributeValue<string>("fullname");
                    return er;
                });
            results.AddRange(contacts);

            // Accounts: emailaddress1
            var accountQ = new QueryExpression("account")
            {
                ColumnSet = new ColumnSet("name", "emailaddress1"),
                TopCount = top
            };
            var accFilter = new FilterExpression(LogicalOperator.Or);
            foreach (var em in emailList)
                accFilter.AddCondition("emailaddress1", ConditionOperator.Equal, em);
            accountQ.Criteria.AddFilter(accFilter);
            var accounts = org.RetrieveMultiple(accountQ).Entities
                .Select(a =>
                {
                    var er = new Microsoft.Xrm.Sdk.EntityReference("account", a.Id);
                    er.Name = a.GetAttributeValue<string>("name");
                    return er;
                });
            results.AddRange(accounts);

            // Leads: emailaddress1
            var leadQ = new QueryExpression("lead")
            {
                ColumnSet = new ColumnSet("fullname", "emailaddress1"),
                TopCount = top
            };
            var leadFilter = new FilterExpression(LogicalOperator.Or);
            foreach (var em in emailList)
                leadFilter.AddCondition("emailaddress1", ConditionOperator.Equal, em);
            leadQ.Criteria.AddFilter(leadFilter);
            var leads = org.RetrieveMultiple(leadQ).Entities
                .Select(l =>
                {
                    var er = new Microsoft.Xrm.Sdk.EntityReference("lead", l.Id);
                    er.Name = l.GetAttributeValue<string>("fullname");
                    return er;
                });
            results.AddRange(leads);

            // De-duplicate by (logicalName, Id)
            return results
                .GroupBy(r => new { r.LogicalName, r.Id })
                .Select(g => g.First())
                .ToList();
        }
    }
}
