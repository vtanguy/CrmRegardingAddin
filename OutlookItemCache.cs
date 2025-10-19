using System;
using System.Collections.Generic;

namespace CrmRegardingAddin
{
    /// <summary>
    /// Ultra-simple in-memory TTL cache (per session) to avoid repeated MAPI reads.
    /// </summary>
    internal static class OutlookItemCache
    {
        private sealed class Entry
        {
            public string Value;
            public DateTime Expire;
        }

        private static readonly Dictionary<string, Entry> _map = new Dictionary<string, Entry>();
        private static readonly object _sync = new object();
        private static readonly TimeSpan _ttl = TimeSpan.FromSeconds(20);

        public static string GetOrAdd(string key, Func<string> factory)
        {
            if (key == null) return null;
            var now = DateTime.UtcNow;

            lock (_sync)
            {
                Entry e;
                if (_map.TryGetValue(key, out e) && e != null && e.Expire > now)
                    return e.Value;
            }

            var v = factory != null ? factory() : null;

            lock (_sync)
            {
                _map[key] = new Entry { Value = v, Expire = DateTime.UtcNow + _ttl };
            }
            return v;
        }
    }
}
