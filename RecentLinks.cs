using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Microsoft.Xrm.Sdk;

namespace CrmRegardingAddin
{
    [Serializable]
    public class RecentLink
    {
        public string EntityLogicalName;
        public Guid Id;
        public string Name;
        public string EntityDisplayName; // friendly display name for the entity (persisted)

        public override string ToString()
        {
            var disp = string.IsNullOrWhiteSpace(EntityDisplayName) ? EntityLogicalName : EntityDisplayName;
            var nm = string.IsNullOrWhiteSpace(Name) ? Id.ToString() : Name;
            return $"[{disp}] {nm}";
        }
    }

    public static class RecentLinks
    {
        private const int MaxItems = 10;
        private static readonly string Folder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CrmRegardingAddin");
        private static readonly string FilePath = Path.Combine(Folder, "recent-links.xml");

        public static List<RecentLink> GetAll()
        {
            try
            {
                if (!File.Exists(FilePath)) return new List<RecentLink>();
                var xs = new XmlSerializer(typeof(List<RecentLink>));
                using (var fs = File.OpenRead(FilePath))
                {
                    var list = (List<RecentLink>)xs.Deserialize(fs);
                    return list ?? new List<RecentLink>();
                }
            }
            catch
            {
                // Back-compat resilience: any issue returns empty list
                return new List<RecentLink>();
            }
        }

        private static void Save(List<RecentLink> items)
        {
            try
            {
                Directory.CreateDirectory(Folder);
                var xs = new XmlSerializer(typeof(List<RecentLink>));
                using (var fs = File.Create(FilePath))
                    xs.Serialize(fs, items);
            }
            catch
            {
                // ignore disk/serialization errors
            }
        }

        public static void Remember(EntityReference er)
        {
            if ( er == null || er.Id == Guid.Empty || string.IsNullOrWhiteSpace(er.LogicalName) )
                return;

            var list = GetAll();

            // de-dup by logical name + id
            list.RemoveAll(x => x != null &&
                                x.Id == er.Id &&
                                !string.IsNullOrWhiteSpace(x.EntityLogicalName) &&
                                x.EntityLogicalName.Equals(er.LogicalName, StringComparison.OrdinalIgnoreCase));

            // Resolve friendly entity display name from search options (best-effort)
            string display = null;
            try
            {
                var opts = CrmSearchService.GetEntityOptions();
                var opt = opts?.FirstOrDefault(o => o != null &&
                    o.LogicalName.Equals(er.LogicalName, StringComparison.OrdinalIgnoreCase));
                display = opt?.DisplayName;
            }
            catch
            {
                // swallow – keep logical name if not available
            }

            list.Insert(0, new RecentLink
            {
                EntityLogicalName = er.LogicalName,
                Id = er.Id,
                Name = string.IsNullOrWhiteSpace(er.Name) ? er.Id.ToString() : er.Name,
                EntityDisplayName = string.IsNullOrWhiteSpace(display) ? er.LogicalName : display
            });

            if (list.Count > MaxItems)
                list = list.GetRange(0, MaxItems);

            Save(list);
        }
    }
}
