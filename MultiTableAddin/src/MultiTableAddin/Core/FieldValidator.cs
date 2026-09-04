using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MultiTableAddin.Core;

/// <summary>字段录入值校验器（支持必填、范围、长度、正则）</summary>
public static class FieldValidator
{
    public static string? Validate(FieldOverride? rule, object? value)
    {
        if (rule == null) return null;

        string text = value?.ToString() ?? string.Empty;
        bool isEmpty = string.IsNullOrWhiteSpace(text);

        if (rule.Required && isEmpty)
        {
            return string.IsNullOrWhiteSpace(rule.ErrorMessage) ? "该字段不能为空" : rule.ErrorMessage;
        }

        if (isEmpty) return null;

        if (rule.MinLength.HasValue && text.Length < rule.MinLength.Value)
        {
            return string.IsNullOrWhiteSpace(rule.ErrorMessage)
                ? $"长度不能少于 {rule.MinLength.Value} 个字符"
                : rule.ErrorMessage;
        }

        if (rule.MaxLength.HasValue && text.Length > rule.MaxLength.Value)
        {
            return string.IsNullOrWhiteSpace(rule.ErrorMessage)
                ? $"长度不能超过 {rule.MaxLength.Value} 个字符"
                : rule.ErrorMessage;
        }

        if (rule.MinValue.HasValue || rule.MaxValue.HasValue)
        {
            if (!ValueFormatter.TryToDouble(value, out double d))
            {
                return string.IsNullOrWhiteSpace(rule.ErrorMessage) ? "请输入有效的数字" : rule.ErrorMessage;
            }

            if (rule.MinValue.HasValue && d < rule.MinValue.Value)
            {
                return string.IsNullOrWhiteSpace(rule.ErrorMessage)
                    ? $"不能小于 {rule.MinValue.Value}"
                    : rule.ErrorMessage;
            }

            if (rule.MaxValue.HasValue && d > rule.MaxValue.Value)
            {
                return string.IsNullOrWhiteSpace(rule.ErrorMessage)
                    ? $"不能大于 {rule.MaxValue.Value}"
                    : rule.ErrorMessage;
            }
        }

        if (!string.IsNullOrWhiteSpace(rule.RegexPattern))
        {
            try
            {
                if (!Regex.IsMatch(text, rule.RegexPattern))
                {
                    return string.IsNullOrWhiteSpace(rule.ErrorMessage)
                        ? "格式不符合要求"
                        : rule.ErrorMessage;
                }
            }
            catch (RegexParseException)
            {
                return $"正则表达式有误: {rule.RegexPattern}";
            }
        }

        return null;
    }

    /// <summary>年龄类专用校验：正整数且大于 0</summary>
    public static string? ValidateAge(object? value)
    {
        string text = value?.ToString() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int age))
            return "年龄必须是整数";
        if (age <= 0) return "年龄必须大于 0";
        if (age > 150) return "年龄不能超过 150";
        return null;
    }

    /// <summary>手机号专用校验：11 位数字</summary>
    public static string? ValidatePhone(object? value)
    {
        string text = value?.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (!Regex.IsMatch(text, @"^\d{11}$"))
            return "手机号必须是 11 位数字";
        return null;
    }
}
