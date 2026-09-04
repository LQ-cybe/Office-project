using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MultiTableAddin.Core;

/// <summary>
/// 字段类型自动识别器。
/// 证据优先级：用户规则库（字段名）&gt; Excel 单元格数字格式 &gt; 实际值样本。
/// 名称规则命中后仍会用数据补齐选项（Select 类型），保证下拉框有真实可选值。
/// </summary>
public static class FieldTypeDetector
{
    private static readonly Regex EmailRegex =
        new(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.Compiled);

    private static readonly Regex PhoneRegex =
        new(@"^(\+?86[-\s]?)?1[3-9]\d{9}$|^0\d{2,3}-?\d{7,8}$", RegexOptions.Compiled);

    private static readonly Regex UrlRegex =
        new(@"^(https?://|www\.)\S+$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// 综合识别一个字段的类型与选项
    /// </summary>
    /// <param name="fieldName">字段名</param>
    /// <param name="numberFormat">Excel 列的 NumberFormat 字符串，可为空</param>
    /// <param name="samples">该列的值样本</param>
    /// <param name="library">规则库，为空时使用全局规则库</param>
    public static FieldSchema Detect(
        string fieldName,
        string? numberFormat,
        IReadOnlyList<object?> samples,
        FieldRuleLibrary? library = null)
    {
        var schema = new FieldSchema
        {
            Name = fieldName,
            NumberFormat = numberFormat ?? string.Empty
        };

        var lib = library ?? FieldRuleLibrary.Current;
        var rule = lib.Match(fieldName);

        if (rule != null)
        {
            schema.Type = rule.Type;
            var ruleOptions = rule.OptionList;
            if (ruleOptions.Count > 0)
            {
                schema.Options = ruleOptions;
            }
            else if (rule.Type == FieldType.Select)
            {
                schema.Options = CollectDistinct(samples, 30);
            }

            // 名称规则说是数字但实际全是文本 → 数据优先，避免录入框类型错乱
            if (FieldTypeHelper.IsNumeric(rule.Type) && samples.Count > 0 && !MostlyNumeric(samples))
            {
                var fallbackType = InferFromValues(samples, out var fallbackOptions);
                schema.Type = fallbackType;
                if (fallbackType == FieldType.Select) schema.Options = fallbackOptions;
            }

            return schema;
        }

        // 没有名称规则时，先看 Excel 数字格式
        var byFormat = InferFromNumberFormat(numberFormat);
        if (byFormat.HasValue)
        {
            schema.Type = byFormat.Value;
            return schema;
        }

        // 最后靠值样本推断
        schema.Type = InferFromValues(samples, out var options);
        if (schema.Type == FieldType.Select) schema.Options = options;
        if (schema.Type == FieldType.Quarter) schema.Options = FieldSchema.QuarterOptions.ToList();
        return schema;
    }

    /// <summary>根据 Excel 单元格数字格式推断类型</summary>
    public static FieldType? InferFromNumberFormat(string? numberFormat)
    {
        if (string.IsNullOrWhiteSpace(numberFormat)) return null;
        string f = numberFormat.Trim();
        if (f is "General" or "@" or "常规") return null;

        // 日期格式：包含 y/m/d 或中文年月日
        bool hasTime = f.Contains('h') || f.Contains('H') || f.Contains("ss");
        bool hasDate = f.Contains('y') || f.Contains("yyyy") || f.Contains('年') ||
                       (f.Contains('m') && f.Contains('d'));
        if (hasDate) return hasTime ? FieldType.DateTime : FieldType.Date;
        if (hasTime) return FieldType.DateTime;

        if (f.Contains('%')) return FieldType.Percentage;
        if (f.Contains('¥') || f.Contains('￥') || f.Contains('$') || f.Contains("#,##0.00_"))
            return FieldType.Currency;
        if (f.Contains("0.0")) return FieldType.Number;
        if (f is "0" || f == "#,##0") return FieldType.Integer;

        return null;
    }

    /// <summary>纯粹依据值样本推断类型</summary>
    public static FieldType InferFromValues(IReadOnlyList<object?> samples, out List<string> options)
    {
        options = new List<string>();

        var valid = samples.Where(v => v != null && !string.IsNullOrWhiteSpace(v.ToString())).ToList();
        if (valid.Count == 0) return FieldType.Text;

        int total = valid.Count;
        int dateCount = 0, numCount = 0, intCount = 0, boolCount = 0;
        int emailCount = 0, phoneCount = 0, urlCount = 0, quarterCount = 0;
        int maxLen = 0;

        foreach (var v in valid)
        {
            string s = v!.ToString()!.Trim();
            maxLen = Math.Max(maxLen, s.Length);

            if (v is bool) { boolCount++; continue; }
            if (v is DateTime) { dateCount++; continue; }

            if (FieldSchema.IsQuarterValue(s)) { quarterCount++; continue; }

            if (ValueFormatter.TryToDouble(v, out double d))
            {
                numCount++;
                if (Math.Abs(d - Math.Round(d)) < 1e-9) intCount++;
                continue;
            }

            if (EmailRegex.IsMatch(s)) { emailCount++; continue; }
            if (PhoneRegex.IsMatch(s)) { phoneCount++; continue; }
            if (UrlRegex.IsMatch(s)) { urlCount++; continue; }
            if (ValueFormatter.TryToDateTime(v, out _)) { dateCount++; }
        }

        double Ratio(int n) => (double)n / total;
        const double Threshold = 0.8;

        if (Ratio(quarterCount) >= Threshold) return FieldType.Quarter;
        if (Ratio(boolCount) >= Threshold) return FieldType.Checkbox;
        if (Ratio(dateCount) >= Threshold) return FieldType.Date;
        if (Ratio(emailCount) >= Threshold) return FieldType.Email;
        if (Ratio(phoneCount) >= Threshold) return FieldType.Phone;
        if (Ratio(urlCount) >= Threshold) return FieldType.Url;

        if (Ratio(numCount) >= Threshold)
            return Ratio(intCount) >= 0.95 ? FieldType.Integer : FieldType.Number;

        // 文本：长文本 / 枚举 / 普通文本
        if (maxLen > 40) return FieldType.LongText;

        var distinct = CollectDistinct(samples, 40);
        // 取值集中且行数足够多时，视为可枚举的单选字段
        if (distinct.Count > 0 && distinct.Count <= 15 && total >= distinct.Count * 2)
        {
            options = distinct;
            return FieldType.Select;
        }

        return FieldType.Text;
    }

    private static bool MostlyNumeric(IReadOnlyList<object?> samples)
    {
        var valid = samples.Where(v => v != null && !string.IsNullOrWhiteSpace(v.ToString())).ToList();
        if (valid.Count == 0) return true; // 空列不推翻名称规则
        int n = valid.Count(v => ValueFormatter.TryToDouble(v, out _));
        return (double)n / valid.Count >= 0.6;
    }

    /// <summary>收集列内不重复的取值</summary>
    public static List<string> CollectDistinct(IReadOnlyList<object?> samples, int max)
    {
        return samples
            .Where(v => v != null && !string.IsNullOrWhiteSpace(v.ToString()))
            .Select(v => ValueFormatter.ToDisplayText(v).Trim())
            .Where(s => s.Length > 0)
            .Distinct()
            .Take(max)
            .ToList();
    }
}
