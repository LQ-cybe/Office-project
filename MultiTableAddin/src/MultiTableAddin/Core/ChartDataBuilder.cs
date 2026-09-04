using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MultiTableAddin.Core;

/// <summary>图表上的一个数据点</summary>
public class ChartPoint
{
    public string Label { get; set; } = string.Empty;
    public double Value { get; set; }

    /// <summary>用于排序的键（时间维度时为区间起点）</summary>
    public DateTime? TimeKey { get; set; }

    /// <summary>该点对应的原始行，供钻取使用</summary>
    public List<DataRowModel> Rows { get; set; } = new();
}

/// <summary>图表的一个系列</summary>
public class ChartSeries
{
    public string Name { get; set; } = string.Empty;
    public List<ChartPoint> Points { get; set; } = new();
    public string Color { get; set; } = "#4E7CF6";
}

/// <summary>图表聚合结果</summary>
public class ChartDataSet
{
    public string Title { get; set; } = string.Empty;
    public List<string> Categories { get; set; } = new();
    public List<ChartSeries> Series { get; set; } = new();
    public double MaxValue { get; set; }
    public double MinValue { get; set; }

    /// <summary>无法绘制时的提示文案</summary>
    public string Message { get; set; } = string.Empty;

    public bool IsEmpty => Series.Count == 0 || Categories.Count == 0;

    /// <summary>所有点合计，饼图算占比用</summary>
    public double Total => Series.SelectMany(s => s.Points).Sum(p => p.Value);
}

/// <summary>
/// 图表数据聚合引擎：原始行 + ChartConfig → ChartDataSet。
/// 纯计算，不依赖 WPF 与 Excel，方便替换渲染层（自绘 / LiveCharts2）。
/// </summary>
public static class ChartDataBuilder
{
    /// <summary>默认配色（与 HandyControl 深色主题协调）</summary>
    public static readonly string[] Palette =
    {
        "#4E7CF6", "#F2994A", "#27AE60", "#9B51E0", "#EB5757",
        "#2D9CDB", "#F2C94C", "#00B8A9", "#FF6B9D", "#6C7A89",
        "#845EC2", "#FF8066"
    };

    public static string ColorAt(int index) => Palette[Math.Abs(index) % Palette.Length];

    // ─────────────────────────────────────────────────────────────
    // 聚合
    // ─────────────────────────────────────────────────────────────

    /// <summary>对一组行按指定字段与聚合方式求值</summary>
    public static double Aggregate(IEnumerable<DataRowModel> rows, string field, AggregateMode mode)
    {
        var list = rows as IList<DataRowModel> ?? rows.ToList();

        if (mode == AggregateMode.Count) return list.Count;

        if (mode == AggregateMode.DistinctCount)
        {
            if (string.IsNullOrEmpty(field)) return list.Count;
            return list.Select(r => ValueFormatter.ToDisplayText(r.GetValue(field)).Trim())
                       .Where(s => s.Length > 0)
                       .Distinct()
                       .Count();
        }

        if (string.IsNullOrEmpty(field)) return list.Count;

        var nums = new List<double>(list.Count);
        foreach (var r in list)
            if (ValueFormatter.TryToDouble(r.GetValue(field), out double v))
                nums.Add(v);

        if (nums.Count == 0) return 0;

        return mode switch
        {
            AggregateMode.Sum => nums.Sum(),
            AggregateMode.Average => nums.Average(),
            AggregateMode.Max => nums.Max(),
            AggregateMode.Min => nums.Min(),
            _ => nums.Count
        };
    }

    // ─────────────────────────────────────────────────────────────
    // 时间分组
    // ─────────────────────────────────────────────────────────────

    /// <summary>时间维度分组键文本</summary>
    public static string TimeLabel(DateTime dt, TimeDimension dim) => dim switch
    {
        TimeDimension.Year => dt.Year + "年",
        TimeDimension.Quarter => dt.Year + " " + QuarterName(dt),
        TimeDimension.Month => dt.ToString("yyyy-MM", CultureInfo.InvariantCulture),
        TimeDimension.Week => dt.Year + " 第" + IsoWeek(dt).ToString("00") + "周",
        TimeDimension.Day => dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        _ => dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
    };

    /// <summary>时间维度分组区间起点，用于排序</summary>
    public static DateTime TimeBucket(DateTime dt, TimeDimension dim) => dim switch
    {
        TimeDimension.Year => new DateTime(dt.Year, 1, 1),
        TimeDimension.Quarter => new DateTime(dt.Year, (Quarter(dt) - 1) * 3 + 1, 1),
        TimeDimension.Month => new DateTime(dt.Year, dt.Month, 1),
        TimeDimension.Week => StartOfWeek(dt),
        TimeDimension.Day => dt.Date,
        _ => dt.Date
    };

    public static int Quarter(DateTime dt) => (dt.Month - 1) / 3 + 1;

    public static string QuarterName(DateTime dt) => Quarter(dt) switch
    {
        1 => "第一季度",
        2 => "第二季度",
        3 => "第三季度",
        _ => "第四季度"
    };

    public static DateTime StartOfWeek(DateTime dt)
    {
        int diff = (7 + (int)dt.DayOfWeek - (int)DayOfWeek.Monday) % 7;
        return dt.Date.AddDays(-diff);
    }

    public static int IsoWeek(DateTime dt) =>
        ISOWeek.GetWeekOfYear(dt);

    // ─────────────────────────────────────────────────────────────
    // 构建
    // ─────────────────────────────────────────────────────────────

    /// <summary>根据配置构建图表数据集</summary>
    public static ChartDataSet Build(IReadOnlyList<DataRowModel> rows, ChartConfig cfg)
    {
        var ds = new ChartDataSet { Title = cfg.Title };

        if (rows == null || rows.Count == 0)
        {
            ds.Message = "当前筛选条件下没有数据";
            return ds;
        }

        bool useTime = cfg.TimeGroup != TimeDimension.None && !string.IsNullOrEmpty(cfg.TimeField);
        if (!useTime && string.IsNullOrEmpty(cfg.DimensionField))
        {
            ds.Message = "请先选择分组维度或时间字段";
            return ds;
        }

        // 仪表盘图是单值，走独立分支
        if (cfg.Type == ChartType.Gauge)
        {
            double val = Aggregate(rows, cfg.MetricField, cfg.Aggregation);
            ds.Categories.Add(cfg.Title);
            ds.Series.Add(new ChartSeries
            {
                Name = cfg.Title,
                Color = ColorAt(0),
                Points = { new ChartPoint { Label = cfg.Title, Value = val, Rows = rows.ToList() } }
            });
            ds.MaxValue = Math.Max(cfg.GaugeTarget, val);
            ds.MinValue = 0;
            return ds;
        }

        // ── 一级分组：类别轴 ────────────────────────────────────
        var buckets = new List<(string Label, DateTime? Key, List<DataRowModel> Rows)>();

        if (useTime)
        {
            var map = new Dictionary<DateTime, (string Label, List<DataRowModel> Rows)>();
            int skipped = 0;
            foreach (var r in rows)
            {
                if (!ValueFormatter.TryToDateTime(r.GetValue(cfg.TimeField), out DateTime dt)) { skipped++; continue; }
                var key = TimeBucket(dt, cfg.TimeGroup);
                if (!map.TryGetValue(key, out var entry))
                {
                    entry = (TimeLabel(dt, cfg.TimeGroup), new List<DataRowModel>());
                    map[key] = entry;
                }
                entry.Rows.Add(r);
            }

            if (map.Count == 0)
            {
                ds.Message = $"字段「{cfg.TimeField}」中没有可识别的日期";
                return ds;
            }

            foreach (var kv in map.OrderBy(k => k.Key))
                buckets.Add((kv.Value.Label, kv.Key, kv.Value.Rows));
        }
        else
        {
            var map = new Dictionary<string, List<DataRowModel>>(StringComparer.Ordinal);
            foreach (var r in rows)
            {
                string key = ValueFormatter.ToDisplayText(r.GetValue(cfg.DimensionField)).Trim();
                if (key.Length == 0) key = "(空)";
                if (!map.TryGetValue(key, out var list))
                {
                    list = new List<DataRowModel>();
                    map[key] = list;
                }
                list.Add(r);
            }

            var ordered = map
                .Select(kv => (Label: kv.Key, Key: (DateTime?)null, Rows: kv.Value))
                .OrderByDescending(t => Aggregate(t.Rows, cfg.MetricField, cfg.Aggregation))
                .ToList();

            // TopN 之外的合并为「其他」
            int topN = cfg.TopN <= 0 ? int.MaxValue : cfg.TopN;
            if (ordered.Count > topN)
            {
                var head = ordered.Take(topN - 1).ToList();
                var tailRows = ordered.Skip(topN - 1).SelectMany(t => t.Rows).ToList();
                head.Add(("其他", null, tailRows));
                ordered = head;
            }

            buckets.AddRange(ordered);
        }

        ds.Categories = buckets.Select(b => b.Label).ToList();

        // ── 二级分组：多系列 ────────────────────────────────────
        bool multiSeries = !string.IsNullOrEmpty(cfg.SeriesField)
                           && cfg.Type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area;

        if (multiSeries)
        {
            var seriesKeys = rows
                .Select(r => ValueFormatter.ToDisplayText(r.GetValue(cfg.SeriesField)).Trim())
                .Select(s => s.Length == 0 ? "(空)" : s)
                .Distinct()
                .Take(8)
                .ToList();

            int ci = 0;
            foreach (var sk in seriesKeys)
            {
                var series = new ChartSeries { Name = sk, Color = ColorAt(ci++) };
                foreach (var b in buckets)
                {
                    var sub = b.Rows.Where(r =>
                    {
                        string v = ValueFormatter.ToDisplayText(r.GetValue(cfg.SeriesField)).Trim();
                        return (v.Length == 0 ? "(空)" : v) == sk;
                    }).ToList();

                    series.Points.Add(new ChartPoint
                    {
                        Label = b.Label,
                        TimeKey = b.Key,
                        Value = sub.Count == 0 ? 0 : Aggregate(sub, cfg.MetricField, cfg.Aggregation),
                        Rows = sub
                    });
                }
                ds.Series.Add(series);
            }
        }
        else
        {
            var series = new ChartSeries
            {
                Name = string.IsNullOrEmpty(cfg.MetricField)
                    ? AggregateModeHelper.GetLabel(cfg.Aggregation)
                    : cfg.MetricField + " " + AggregateModeHelper.GetLabel(cfg.Aggregation),
                Color = ColorAt(0)
            };

            int ci = 0;
            foreach (var b in buckets)
            {
                series.Points.Add(new ChartPoint
                {
                    Label = b.Label,
                    TimeKey = b.Key,
                    Value = Aggregate(b.Rows, cfg.MetricField, cfg.Aggregation),
                    Rows = b.Rows
                });
                ci++;
            }
            ds.Series.Add(series);
        }

        var allValues = ds.Series.SelectMany(s => s.Points).Select(p => p.Value).ToList();
        ds.MaxValue = allValues.Count == 0 ? 0 : allValues.Max();
        ds.MinValue = allValues.Count == 0 ? 0 : Math.Min(0, allValues.Min());

        if (Math.Abs(ds.MaxValue) < 1e-9 && Math.Abs(ds.MinValue) < 1e-9)
            ds.Message = "聚合结果全为 0，请检查度量字段是否为数值列";

        return ds;
    }

    /// <summary>KPI 卡片取值</summary>
    public static double BuildStat(IReadOnlyList<DataRowModel> rows, StatCardConfig cfg)
    {
        IEnumerable<DataRowModel> src = rows;
        if (!string.IsNullOrWhiteSpace(cfg.Filter))
            src = rows.Where(r => ViewEngine.Match(r, cfg.Filter));
        return Aggregate(src, cfg.Field, cfg.Aggregation);
    }
}
