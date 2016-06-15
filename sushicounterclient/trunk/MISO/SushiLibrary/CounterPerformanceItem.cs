/*
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace SushiLibrary
{
    public class CounterPerformanceItem
    {
        public string YOP { get; set; }

        public Dictionary<string, CounterMetric> Metrics
        {
            get
            {
                return _metrics;
            }
        }

        private readonly Dictionary<string, CounterMetric> _metrics = new Dictionary<string, CounterMetric>();

        private static string CalculateHashKey(DateTime start, DateTime end, CounterMetricCategory category)
        {
            return start + "|" + end + "|" + category;
        }

        public CounterMetric GetMetric(DateTime start, DateTime end, CounterMetricCategory category)
        {
            string key = CalculateHashKey(start, end, category);
            if (!_metrics.ContainsKey(key))
            {
                _metrics[key] = new CounterMetric(start, end, category);
            }

            return _metrics[key];
        }

        public bool TryGetMetric(DateTime start, DateTime end, CounterMetricCategory category, out CounterMetric metric)
        {
            string key = CalculateHashKey(start, end, category);
            return _metrics.TryGetValue(key, out metric);
        }
    }
}
