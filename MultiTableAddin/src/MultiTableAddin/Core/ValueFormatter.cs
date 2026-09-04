using System;
using System.Globalization;

namespace MultiTableAddin.Core;

/// <summary>
/// 统一的值格式化 / 解析工具。
/// 视图层展示、写回 Excel、图表聚合都走这里，避免各处 ToString 行为不一致。
/// </summary>
public static class ValueFormatter
{
    /// <summary>把任意值转换为界面展示文本</summary>
    public static string ToDisplayText(object? value)
    {
        if (value == null) return string.Empty;

        return value switch
        {
            DateTime dt => dt.TimeOfDay == TimeSpan.Zero
                ? dt.ToString("yyyy-MM-dd")
                : dt.ToString("yyyy-MM-dd HH:mm"),
            bool b => b ? "是" : "否",
            double d => FormatDouble(d),
            float f => FormatDouble(f),
            decimal m => FormatDouble((double)m),
            _ => value.ToString() ?? string.Empty
        };
    }

    /// <summary>按字段类型格式化展示文本</summary>
    public static string ToDisplayText(object? value, FieldType type)
    {
        if (value == null) return string.Empty;

        switch (type)
        {
            case FieldType.Currency:
                if (TryToDouble(value, out double money))
                    return "¥" + money.ToString("#,##0.00", CultureInfo.InvariantCulture);
                break;

            case FieldType.Percentage:
                if (TryToDouble(value, out double pct))
                {
                    // Excel 中百分比单元格 Value2 为 0.85 这种小数
                    double shown = Math.Abs(pct) <= 1.000001 ? pct * 100 : pct;
                    return shown.ToString("0.##", CultureInfo.InvariantCulture) + "%";
                }
                break;

            case FieldType.Integer:
                if (TryToDouble(value, out double iv))
                    return Math.Round(iv).ToString("#,##0", CultureInfo.InvariantCulture);
                break;

            case FieldType.Date:
                if (TryToDateTime(value, out DateTime dv)) return dv.ToString("yyyy-MM-dd");
                break;

            case FieldType.DateTime:
                if (TryToDateTime(value, out DateTime dtv)) return dtv.ToString("yyyy-MM-dd HH:mm");
                break;

            case FieldType.Quarter:
                return FieldSchema.NormalizeQuarter(value.ToString() ?? string.Empty);

            case FieldType.Checkbox:
                if (value is bool cb) return cb ? "是" : "否";
                break;
        }

        return ToDisplayText(value);
    }

    private static string FormatDouble(double d)
    {
        if (Math.Abs(d - Math.Round(d)) < 1e-10 && Math.Abs(d) < 1e15)
            return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
        return d.ToString("0.####", CultureInfo.InvariantCulture);
    }

    /// <summary>尝试把值转为 double</summary>
    public static bool TryToDouble(object? value, out double result)
    {
        result = 0;
        switch (value)
        {
            case null:
                return false;
            case double d:
                result = d; return true;
            case int i:
                result = i; return true;
            case long l:
                result = l; return true;
            case float f:
                result = f; return true;
            case decimal m:
                result = (double)m; return true;
            case bool b:
                result = b ? 1 : 0; return true;
            case DateTime dt:
                result = dt.ToOADate(); return true;
        }

        string s = value.ToString()?.Trim() ?? string.Empty;
        if (s.Length == 0) return false;

        // 去掉常见货币符号、千分位、百分号
        bool isPercent = s.EndsWith("%", StringComparison.Ordinal);
        s = s.Replace("¥", string.Empty)
             .Replace("￥", string.Empty)
             .Replace("$", string.Empty)
             .Replace(",", string.Empty)
             .Replace("，", string.Empty)
             .Replace("%", string.Empty)
             .Trim();

        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
        {
            if (isPercent) result /= 100.0;
            return true;
        }
        return false;
    }

    /// <summary>尝试把值转为 DateTime</summary>
    public static bool TryToDateTime(object? value, out DateTime result)
    {
        result = default;
        switch (value)
        {
            case null:
                return false;
            case DateTime dt:
                result = dt; return true;
            case double d:
                return TryFromOADate(d, out result);
            case int i:
                return TryFromOADate(i, out result);
        }

        string s = value.ToString()?.Trim() ?? string.Empty;
        if (s.Length == 0) return false;

        if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out result)) return true;
        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out result)) return true;

        string[] formats =
        {
            "yyyy-MM-dd", "yyyy/M/d", "yyyy.M.d", "yyyyMMdd",
            "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
            "yyyy年M月d日", "M/d/yyyy", "d-MMM-yyyy"
        };
        return DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
    }

    /// <summary>Excel OLE 日期序列号转 DateTime（带合理区间校验）</summary>
    public static bool TryFromOADate(double serial, out DateTime result)
    {
        result = default;
        // 1900-01-01 => 1, 2199-12-31 => 109573；限定区间避免把普通数字误判为日期
        if (serial < 1 || serial > 109573) return false;
        try
        {
            result = DateTime.FromOADate(serial);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>按字段类型把界面输入解析为存储值</summary>
    public static object? ParseInput(string? text, FieldType type)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        text = text.Trim();

        switch (type)
        {
            case FieldType.Number:
            case FieldType.Currency:
                return TryToDouble(text, out double d) ? d : text;

            case FieldType.Integer:
                return TryToDouble(text, out double i) ? (object)(long)Math.Round(i) : text;

            case FieldType.Percentage:
                if (TryToDouble(text, out double p))
                {
                    // 用户输入 85 表示 85%
                    return text.EndsWith("%", StringComparison.Ordinal) || Math.Abs(p) <= 1.000001
                        ? p
                        : p / 100.0;
                }
                return text;

            case FieldType.Date:
            case FieldType.DateTime:
                return TryToDateTime(text, out DateTime dt) ? dt : text;

            case FieldType.Checkbox:
                return text is "是" or "true" or "TRUE" or "1" or "√";

            case FieldType.Quarter:
                return FieldSchema.NormalizeQuarter(text);

            default:
                return text;
        }
    }

    /// <summary>把数字格式化为紧凑的 KPI 展示文本（万 / 亿）</summary>
    public static string ToCompactNumber(double value, string format = "auto")
    {
        switch (format)
        {
            case "int":
                return Math.Round(value).ToString("#,##0", CultureInfo.InvariantCulture);
            case "percent":
                return (Math.Abs(value) <= 1.000001 ? value * 100 : value)
                    .ToString("0.#", CultureInfo.InvariantCulture) + "%";
            case "money":
                return "¥" + CompactCore(value);
        }

        return CompactCore(value);
    }

    private static string CompactCore(double value)
    {
        double abs = Math.Abs(value);
        if (abs >= 100000000) return (value / 100000000).ToString("0.##", CultureInfo.InvariantCulture) + "亿";
        if (abs >= 10000) return (value / 10000).ToString("0.##", CultureInfo.InvariantCulture) + "万";
        if (Math.Abs(value - Math.Round(value)) < 1e-9)
            return ((long)Math.Round(value)).ToString("#,##0", CultureInfo.InvariantCulture);
        return value.ToString("#,##0.##", CultureInfo.InvariantCulture);
    }

    /// <summary>坐标轴刻度文本</summary>
    public static string ToAxisLabel(double value)
    {
        double abs = Math.Abs(value);
        if (abs >= 100000000) return (value / 100000000).ToString("0.#", CultureInfo.InvariantCulture) + "亿";
        if (abs >= 10000) return (value / 10000).ToString("0.#", CultureInfo.InvariantCulture) + "万";
        if (abs >= 1000) return value.ToString("#,##0", CultureInfo.InvariantCulture);
        if (Math.Abs(value - Math.Round(value)) < 1e-9) return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        return value.ToString("0.##", CultureInfo.InvariantCulture);
    }
}
