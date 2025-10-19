using System;
using System.Collections.Generic;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    /// <summary>
    /// Soft limiter for MAPI PropertyAccessor calls to avoid Exchange throttling.
    /// Returns null when the budget is exceeded instead of blocking the UI thread.
    /// </summary>
    internal static class MapiLimiter
    {
        // Tunables (kept simple for VS2015 / C#6)
        private static readonly int _maxCallsPerSecond = 20;     // default budget
        private static readonly int _maxBurst = 30;               // small burst tolerance
        private static readonly TimeSpan _window = TimeSpan.FromSeconds(1);

        private static readonly object _sync = new object();
        private static readonly Queue<DateTime> _calls = new Queue<DateTime>();

        public static bool Enabled = true;

        public static int ApproxCallsInWindow
        {
            get
            {
                lock (_sync)
                {
                    TrimLocked();
                    return _calls.Count;
                }
            }
        }

        private static void TrimLocked()
        {
            var limit = DateTime.UtcNow - _window;
            while (_calls.Count > 0 && _calls.Peek() < limit)
                _calls.Dequeue();
        }

        private static bool TryEnterBudget()
        {
            lock (_sync)
            {
                TrimLocked();
                if (_calls.Count >= Math.Max(_maxCallsPerSecond, _maxBurst))
                    return FalseFast();

                _calls.Enqueue(DateTime.UtcNow);
                return true;
            }
        }

        private static bool FalseFast() { return false; }

        public static object TryGetProperty(Outlook.PropertyAccessor pa, string dasl)
        {
            if (!Enabled || pa == null || string.IsNullOrEmpty(dasl))
                return SafeGet(pa, dasl);

            if (!TryEnterBudget())
                return null;
            return SafeGet(pa, dasl);
        }

        public static string TryGetString(Outlook.PropertyAccessor pa, string dasl)
        {
            var o = TryGetProperty(pa, dasl);
            return o as string;
        }

        public static double? TryGetDouble(Outlook.PropertyAccessor pa, string dasl)
        {
            var o = TryGetProperty(pa, dasl);
            if (o == null) return null;
            try
            {
                if (o is double) return (double)o;
                if (o is float) return (double)(float)o;
                if (o is int) return (double)(int)o;
                double d;
                if (o is string && double.TryParse((string)o, out d)) return d;
            }
            catch { }
            return null;
        }

        private static object SafeGet(Outlook.PropertyAccessor pa, string dasl)
        {
            try { return pa != null ? pa.GetProperty(dasl) : null; }
            catch { return null; }
        }
    }
}
