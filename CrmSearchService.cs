using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace CrmRegardingAddin
{
    public class SearchEntityOption
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public override string ToString() { return DisplayName; }
    }

    public class ViewInfo
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string EntityName { get; set; }
        public bool IsQuickFind { get; set; }
        public bool IsPersonal { get; set; }
        public string FetchXml { get; set; }
        public string LayoutXml { get; set; }
        public override string ToString() { return Name + (IsQuickFind ? " (Recherche rapide)" : "") + (IsPersonal ? " (Perso)" : ""); }
    }

    public class ColumnSpec
    {
        public string EntityLogicalName; // entité de la colonne (main par défaut)
        public string Attribute;         // logical name
        public int Width;
        public string Header;            // DisplayName (rempli via métadonnées)
    }

    internal static class MetadataCache
    {
        private static readonly Dictionary<string, EntityMetadata> _entity = new Dictionary<string, EntityMetadata>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _gate = new object();

        public static EntityMetadata GetEntity(IOrganizationService org, string entityLogicalName)
        {
            if (string.IsNullOrWhiteSpace(entityLogicalName)) return null;
            EntityMetadata meta;
            if (_entity.TryGetValue(entityLogicalName, out meta)) return meta;

            lock (_gate)
            {
                if (_entity.TryGetValue(entityLogicalName, out meta)) return meta;
                var req = new RetrieveEntityRequest
                {
                    LogicalName = entityLogicalName,
                    EntityFilters = EntityFilters.Attributes | EntityFilters.Entity
                };
                var resp = (RetrieveEntityResponse)org.Execute(req);
                meta = resp.EntityMetadata;
                _entity[entityLogicalName] = meta;
                return meta;
            }
        }

        public static string GetAttributeDisplayName(IOrganizationService org, string entityLogicalName, string attributeLogicalName)
        {
            var meta = GetEntity(org, entityLogicalName);
            if (meta == null) return attributeLogicalName;
            var a = meta.Attributes.FirstOrDefault(x => string.Equals(x.LogicalName, attributeLogicalName, StringComparison.OrdinalIgnoreCase));
            if (a == null || a.DisplayName == null) return attributeLogicalName;
            var ll = a.DisplayName.UserLocalizedLabel ?? a.DisplayName.LocalizedLabels.FirstOrDefault();
            return ll != null && !string.IsNullOrWhiteSpace(ll.Label) ? ll.Label : attributeLogicalName;
        }

        public static string GetOptionLabel(IOrganizationService org, string entityLogicalName, string attributeLogicalName, int value)
        {
            var meta = GetEntity(org, entityLogicalName);
            if (meta == null) return value.ToString();

            var a = meta.Attributes.FirstOrDefault(x => string.Equals(x.LogicalName, attributeLogicalName, StringComparison.OrdinalIgnoreCase));
            if (a == null) return value.ToString();

            // Status / State
            var status = a as StatusAttributeMetadata;
            if (status != null)
            {
                var opt = status.OptionSet.Options.OfType<StatusOptionMetadata>().FirstOrDefault(o => (o.Value ?? -1) == value);
                if (opt != null) return GetLabel(opt.Label);
            }
            var state = a as StateAttributeMetadata;
            if (state != null)
            {
                var opt = state.OptionSet.Options.OfType<StateOptionMetadata>().FirstOrDefault(o => (o.Value ?? -1) == value);
                if (opt != null) return GetLabel(opt.Label);
            }

            // Picklist
            var pick = a as PicklistAttributeMetadata;
            if (pick != null && pick.OptionSet != null)
            {
                var opt = pick.OptionSet.Options.FirstOrDefault(o => (o.Value ?? -1) == value);
                if (opt != null) return GetLabel(opt.Label);
            }

            return value.ToString();
        }

        public static string GetBooleanLabel(IOrganizationService org, string entityLogicalName, string attributeLogicalName, bool value)
        {
            var meta = GetEntity(org, entityLogicalName);
            if (meta == null) return value ? "Oui" : "Non";
            var a = meta.Attributes.FirstOrDefault(x => string.Equals(x.LogicalName, attributeLogicalName, StringComparison.OrdinalIgnoreCase));
            var b = a as BooleanAttributeMetadata;
            if (b == null) return value ? "Oui" : "Non";
            return value ? GetLabel(b.OptionSet.TrueOption.Label) : GetLabel(b.OptionSet.FalseOption.Label);
        }

        private static string GetLabel(Label label)
        {
            if (label == null) return "";
            var l = label.UserLocalizedLabel ?? label.LocalizedLabels.FirstOrDefault();
            return l != null ? l.Label : "";
        }
    }

    public static class CrmSearchService
    {
        public static List<ColumnSpec> ParseLayoutColumns(IOrganizationService org, string entityLogicalName, string layoutXml)
        {
            var cols = new List<ColumnSpec>();
            if (string.IsNullOrWhiteSpace(layoutXml))
                return cols;

            try
            {
                var doc = XDocument.Parse(layoutXml);
                var root = doc.Root;
                if (root == null) return cols;
                XNamespace ns = root.Name.Namespace;

                // CRM layout: <grid><row><cell name="xxx" width="###" ... /></row></grid>
                var cells = doc.Descendants(ns + "cell");
                foreach (var cell in cells)
                {
                    var nameAttr = cell.Attribute("name");
                    if (nameAttr == null) continue;
                    var logicalName = nameAttr.Value ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(logicalName)) continue;

                    int width = 150;
                    var wAttr = cell.Attribute("width");
                    if (wAttr != null)
                    {
                        int w;
                        if (int.TryParse(wAttr.Value, out w)) width = w;
                    }

                    // Header (display name)
                    string headerAttrLogical = logicalName;
                    // alias "entity.attribute"
                    if (headerAttrLogical.Contains("."))
                    {
                        var parts = headerAttrLogical.Split('.');
                        if (parts.Length == 2) headerAttrLogical = parts[1];
                    }

                    // Strip '*name' suffix (lookup display)
                    string labelAttr = headerAttrLogical.EndsWith("name", StringComparison.OrdinalIgnoreCase)
                        ? headerAttrLogical.Substring(0, headerAttrLogical.Length - 4)
                        : headerAttrLogical;

                    string header = MetadataCache.GetAttributeDisplayName(org, entityLogicalName, labelAttr);
                    if (string.IsNullOrWhiteSpace(header)) header = headerAttrLogical; // fallback

                    cols.Add(new ColumnSpec
                    {
                        EntityLogicalName = entityLogicalName,
                        Attribute = logicalName,
                        Width = width,
                        Header = header
                    });
                }
            }
            catch
            {
                // ignore parse errors, return what we have
            }
            return cols;
        }

        public static List<SearchEntityOption> GetEntityOptions()
        {
            return new List<SearchEntityOption>
            {
                new SearchEntityOption { LogicalName = "contact",      DisplayName = "Contacts"  }, // défaut
                new SearchEntityOption { LogicalName = "account",      DisplayName = "Sociétés"  },
                new SearchEntityOption { LogicalName = "new_instance", DisplayName = "Instances" }
            };
        }

        public static List<ViewInfo> GetViews(IOrganizationService org, string entityLogicalName)
        {
            var views = new List<ViewInfo>();

            // Vues système
            var qSys = new QueryExpression("savedquery")
            {
                ColumnSet = new ColumnSet("savedqueryid", "name", "returnedtypecode", "querytype", "fetchxml", "layoutxml", "isdefault"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };
            qSys.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, entityLogicalName);
            qSys.Criteria.AddCondition("querytype", ConditionOperator.Equal, 0);
            qSys.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            var sysRes = org.RetrieveMultiple(qSys);
            foreach (var e in sysRes.Entities)
            {
                views.Add(new ViewInfo
                {
                    Id = e.Id,
                    Name = e.GetAttributeValue<string>("name"),
                    EntityName = entityLogicalName,
                    IsQuickFind = false,
                    IsPersonal = false,
                    FetchXml = e.GetAttributeValue<string>("fetchxml"),
                    LayoutXml = e.GetAttributeValue<string>("layoutxml")
                });
            }

            // Quick Find
            var qQf = new QueryExpression("savedquery")
            {
                ColumnSet = new ColumnSet("savedqueryid", "name", "returnedtypecode", "querytype", "fetchxml", "layoutxml"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };
            qQf.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, entityLogicalName);
            qQf.Criteria.AddCondition("querytype", ConditionOperator.Equal, 4);
            qQf.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            var qf = org.RetrieveMultiple(qQf).Entities.FirstOrDefault();
            if (qf != null)
            {
                views.Insert(0, new ViewInfo
                {
                    Id = qf.Id,
                    Name = "Recherche rapide",
                    EntityName = entityLogicalName,
                    IsQuickFind = true,
                    IsPersonal = false,
                    FetchXml = qf.GetAttributeValue<string>("fetchxml"),
                    LayoutXml = qf.GetAttributeValue<string>("layoutxml")
                });
            }

            // Vues perso
            var qPerso = new QueryExpression("userquery")
            {
                ColumnSet = new ColumnSet("userqueryid", "name", "returnedtypecode", "fetchxml", "layoutxml"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };
            qPerso.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, entityLogicalName);
            qPerso.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            var upRes = org.RetrieveMultiple(qPerso);
            foreach (var e in upRes.Entities)
            {
                views.Add(new ViewInfo
                {
                    Id = e.Id,
                    Name = e.GetAttributeValue<string>("name"),
                    EntityName = entityLogicalName,
                    IsQuickFind = false,
                    IsPersonal = true,
                    FetchXml = e.GetAttributeValue<string>("fetchxml"),
                    LayoutXml = e.GetAttributeValue<string>("layoutxml")
                });
            }

            // Vue système par défaut juste après Quick Find
            var def = sysRes.Entities.FirstOrDefault(x => x.GetAttributeValue<bool?>("isdefault") == true);
            if (def != null)
            {
                var dv = views.FirstOrDefault(v => v.Id == def.Id);
                if (dv != null)
                {
                    views.Remove(dv);
                    int insertPos = views.Count > 0 && views[0].IsQuickFind ? 1 : 0;
                    views.Insert(insertPos, dv);
                }
            }

            return views;
        }

        // --- Quick Find columns ---

        public static string[] GetQuickFindAttributes(IOrganizationService org, string entityLogicalName)
        {
            if (org == null || string.IsNullOrWhiteSpace(entityLogicalName))
                return new string[0];

            // 1) depuis la QuickFind: conditions value="{0}" sur LIKE/begins/ends
            try
            {
                var qQf = new QueryExpression("savedquery")
                {
                    ColumnSet = new ColumnSet("savedqueryid", "fetchxml"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                qQf.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, entityLogicalName);
                qQf.Criteria.AddCondition("querytype", ConditionOperator.Equal, 4);
                qQf.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                var qf = org.RetrieveMultiple(qQf).Entities.FirstOrDefault();
                if (qf != null)
                {
                    var fx = qf.GetAttributeValue<string>("fetchxml");
                    var fromQf = ParseQuickFindAttributesFromFetchXml(fx);
                    var normalized = FilterSearchableAttributes(org, entityLogicalName, fromQf);
                    if (normalized.Length > 0) return normalized;
                }
            }
            catch { }

            // 2) Métadonnées QuickFindAttributeNames
            try
            {
                var meta = MetadataCache.GetEntity(org, entityLogicalName);
                var prop = meta != null ? meta.GetType().GetProperty("QuickFindAttributeNames") : null;
                if (prop != null)
                {
                    var arr = prop.GetValue(meta, null) as IEnumerable<string>;
                    if (arr != null)
                    {
                        var normalized = FilterSearchableAttributes(org, entityLogicalName, arr);
                        if (normalized.Length > 0) return normalized;
                    }
                }
            }
            catch { }

            // 3) Fallback
            try
            {
                var meta = MetadataCache.GetEntity(org, entityLogicalName);
                var candidates = new List<string>();
                if (meta != null && !string.IsNullOrEmpty(meta.PrimaryNameAttribute))
                    candidates.Add(meta.PrimaryNameAttribute);

                if (meta != null)
                {
                    foreach (var a in meta.Attributes)
                    {
                        var sa = a as StringAttributeMetadata;
                        if (sa != null &&
                            sa.IsValidForAdvancedFind != null && sa.IsValidForAdvancedFind.Value)
                        {
                            var ln = sa.LogicalName ?? "";
                            if (ln.Contains("email") || ln.Contains("name") || ln.EndsWith("number"))
                                candidates.Add(ln);
                        }
                    }
                }
                return candidates.Distinct().ToArray();
            }
            catch { }

            return new string[0];
        }

        private static string[] ParseQuickFindAttributesFromFetchXml(string fetchXml)
        {
            if (string.IsNullOrWhiteSpace(fetchXml)) return new string[0];
            try
            {
                var doc = XDocument.Parse(fetchXml);
                var root = doc.Root;
                if (root == null) return new string[0];
                var ns = root.Name.Namespace;

                var attrs = new List<string>();
                foreach (var cond in doc.Descendants(ns + "condition"))
                {
                    var valAttr = cond.Attribute("value");
                    var nameAttr = cond.Attribute("attribute");
                    var opAttr = cond.Attribute("operator");
                    if (valAttr == null || nameAttr == null || opAttr == null) continue;

                    var val = valAttr.Value ?? "";
                    var op = opAttr.Value ?? "";
                    if (val.Contains("{0}") &&
                        (op.Equals("like", StringComparison.OrdinalIgnoreCase)
                         || op.Equals("begins-with", StringComparison.OrdinalIgnoreCase)
                         || op.Equals("ends-with", StringComparison.OrdinalIgnoreCase)))
                    {
                        var an = nameAttr.Value;
                        if (!string.IsNullOrWhiteSpace(an))
                            attrs.Add(an);
                    }
                }
                return attrs.Distinct().ToArray();
            }
            catch
            {
                return new string[0];
            }
        }

        /// <summary>
        /// Conserve colonnes texte; Lookup/Customer/Owner -> 'xxxname'
        /// </summary>
        private static string[] FilterSearchableAttributes(IOrganizationService org, string entityLogicalName, IEnumerable<string> attributes)
        {
            if (attributes == null) return new string[0];
            var list = new List<string>();
            var meta = MetadataCache.GetEntity(org, entityLogicalName);
            foreach (var name in attributes.Distinct())
            {
                if (name.EndsWith("name", StringComparison.OrdinalIgnoreCase))
                {
                    list.Add(name);
                    continue;
                }
                var a = meta != null ? meta.Attributes.FirstOrDefault(x => string.Equals(x.LogicalName, name, StringComparison.OrdinalIgnoreCase)) : null;
                if (a == null) continue;

                if (a is StringAttributeMetadata || a is MemoAttributeMetadata)
                {
                    list.Add(name);
                }
                else if (a.AttributeType == AttributeTypeCode.Lookup
                      || a.AttributeType == AttributeTypeCode.Customer
                      || a.AttributeType == AttributeTypeCode.Owner)
                {
                    list.Add(name + "name");
                }
            }
            return list.Distinct().ToArray();
        }

        private static string NormalizePattern(string term)
        {
            if (string.IsNullOrWhiteSpace(term)) return null;
            var t = term.Trim().Replace('*', '%');
            if (!t.Contains("%")) t = "%" + t + "%";
            return t;
        }

        public static string InjectSearchIntoFetchXml(string fetchXml, string[] findAttributes, string term)
        {
            if (string.IsNullOrWhiteSpace(fetchXml) || findAttributes == null || findAttributes.Length == 0 || string.IsNullOrWhiteSpace(term))
                return fetchXml;

            var pattern = NormalizePattern(term);

            var doc = XDocument.Parse(fetchXml);
            var root = doc.Root;
            if (root == null) return fetchXml;

            XNamespace ns = root.Name.Namespace;
            var entityEl = doc.Descendants(ns + "entity").FirstOrDefault();
            if (entityEl == null) return fetchXml;

            var andFilter = new XElement(ns + "filter", new XAttribute("type", "and"));
            var orFilter = new XElement(ns + "filter", new XAttribute("type", "or"));

            foreach (var a in findAttributes)
            {
                orFilter.Add(new XElement(ns + "condition",
                    new XAttribute("attribute", a),
                    new XAttribute("operator", "like"),
                    new XAttribute("value", pattern)));
            }

            andFilter.Add(orFilter);
            entityEl.Add(andFilter);
            return doc.ToString(SaveOptions.DisableFormatting);
        }

        public static EntityCollection Search(IOrganizationService org, string entityLogicalName, ViewInfo view, string term, int top = 1000)
        {
            var findAttrs = GetQuickFindAttributes(org, entityLogicalName);
            var pattern = NormalizePattern(term);

            if (view != null && view.IsQuickFind)
            {
                // Quick Find via FetchXML (supporte xxxname)
                var fx = BuildQuickFindFetch(entityLogicalName, findAttrs, pattern, top);
                return org.RetrieveMultiple(new FetchExpression(fx));
            }
            else
            {
                var baseFetch = (view != null && !string.IsNullOrEmpty(view.FetchXml))
                                ? view.FetchXml
                                : string.Format("<fetch version='1.0'><entity name='{0}'><all-attributes /></entity></fetch>", entityLogicalName);

                var fetch = !string.IsNullOrWhiteSpace(pattern)
                            ? InjectSearchIntoFetchXml(baseFetch, findAttrs, term)
                            : baseFetch;

                var doc = XDocument.Parse(fetch);
                var fetchEl = doc.Root;
                if (fetchEl != null && top > 0)
                {
                    var topAttr = fetchEl.Attribute("top");
                    if (topAttr == null) fetchEl.Add(new XAttribute("top", top));
                    else topAttr.Value = top.ToString();
                }
                fetch = doc.ToString(SaveOptions.DisableFormatting);

                return org.RetrieveMultiple(new FetchExpression(fetch));
            }
        }

        private static string BuildQuickFindFetch(string entityLogicalName, string[] findAttributes, string pattern, int top)
        {
            if (findAttributes == null) findAttributes = new string[0];
            var fx = new XDocument(
                new XElement("fetch",
                    new XAttribute("version", "1.0"),
                    top > 0 ? new XAttribute("top", top) : null,
                    new XElement("entity",
                        new XAttribute("name", entityLogicalName),
                        new XElement("filter",
                            new XAttribute("type", "or"),
                            findAttributes.Select(a =>
                                new XElement("condition",
                                    new XAttribute("attribute", a),
                                    new XAttribute("operator", "like"),
                                    new XAttribute("value", string.IsNullOrWhiteSpace(pattern) ? "%" : pattern)
                                )
                            )
                        ),
                        // Ordonner par les plus récents en premier pour éviter de tronquer les dernières lignes
                        new XElement("order",
                            new XAttribute("attribute", "modifiedon"),
                            new XAttribute("descending", "true")
                        )
                    )
                )
            );
            return fx.ToString(SaveOptions.DisableFormatting);
        }
    }
}
