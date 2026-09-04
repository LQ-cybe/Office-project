using System;
using System.Collections.Generic;
using System.Linq;

namespace MultiTableAddin.Core;

/// <summary>
/// 视图引擎：接收原始数据 + 视图配置 → 输出过滤/排序/分组后的视图数据集
/// 纯逻辑模块，不依赖 Excel COM，不修改原始数据
/// </summary>
public class ViewEngine
{
    /// <summary>
    /// 执行视图变换：过滤 → 排序 → 分组
    /// </summary>
    public ViewDataSet Apply(DataTableModel source, ViewConfig config)
    {
        var rows = source.Rows.AsEnumerable();

        // 1. 过滤
        if (!string.IsNullOrWhiteSpace(config.Filter))
        {
            rows = rows.Where(r => MatchFilter(r, config.Filter));
        }

        // 2. 排序（使用容错比较器，避免同列混合类型时抛异常）
        if (config.Sort != null && config.Sort.Count > 0)
        {
            IOrderedEnumerable<DataRowModel>? ordered = null;
            foreach (var sort in config.Sort)
            {
                if (string.IsNullOrWhiteSpace(sort.Field)) continue;
                var field = sort.Field;
                bool desc = sort.Order?.Equals("desc", StringComparison.OrdinalIgnoreCase) == true;

                if (ordered == null)
                {
                    ordered = desc
                        ? rows.OrderByDescending(r => r.GetValue(field), CellComparer.Instance)
                        : rows.OrderBy(r => r.GetValue(field), CellComparer.Instance);
                }
                else
                {
                    ordered = desc
                        ? ordered.ThenByDescending(r => r.GetValue(field), CellComparer.Instance)
                        : ordered.ThenBy(r => r.GetValue(field), CellComparer.Instance);
                }
            }
            if (ordered != null) rows = ordered;
        }

        var resultRows = rows.ToList();

        // 3. 分组（看板视图用）
        var groups = new List<ViewGroup>();
        if (!string.IsNullOrWhiteSpace(config.GroupBy))
        {
            var groupField = config.GroupBy;
            var grouped = resultRows
                .GroupBy(r => r.GetValue(groupField)?.ToString() ?? "(无分组)")
                .OrderBy(g => g.Key);

            foreach (var g in grouped)
            {
                groups.Add(new ViewGroup
                {
                    Key = g.Key,
                    Rows = g.ToList()
                });
            }
        }
        else
        {
            groups.Add(new ViewGroup
            {
                Key = "全部",
                Rows = resultRows
            });
        }

        return new ViewDataSet
        {
            Groups = groups,
            TotalCount = resultRows.Count
        };
    }

    /// <summary>对外暴露的筛选判定，供 KPI 卡片等局部统计复用</summary>
    public static bool Match(DataRowModel row, string? filter)
    {
        if (string.IsNullOrWhiteSpace(filter)) return true;
        return MatchFilter(row, filter!);
    }

    /// <summary>简单筛选表达式匹配</summary>
    private static bool MatchFilter(DataRowModel row, string filter)
    {
        // 支持: [字段名] = '值', [字段名] != '值', [字段名] != ''
        // 支持 AND / OR 组合
        try
        {
            return EvaluateExpression(row, filter);
        }
        catch
        {
            return true; // 解析失败则不过滤
        }
    }

    private static bool EvaluateExpression(DataRowModel row, string expr)
    {
        expr = expr.Trim();

        // 处理 OR
        var orParts = SplitByOperator(expr, " OR ", StringComparison.OrdinalIgnoreCase);
        if (orParts.Count > 1)
        {
            return orParts.Any(p => EvaluateExpression(row, p));
        }

        // 处理 AND
        var andParts = SplitByOperator(expr, " AND ", StringComparison.OrdinalIgnoreCase);
        if (andParts.Count > 1)
        {
            return andParts.All(p => EvaluateExpression(row, p));
        }

        // 单个条件: [字段名] op '值'
        return EvaluateCondition(row, expr);
    }

    private static bool EvaluateCondition(DataRowModel row, string condition)
    {
        condition = condition.Trim();

        // 解析 [字段名] 操作符 值
        // 支持: =, !=, >=, <=, >, <
        string[] ops = { "!=", ">=", "<=", "=", ">", "<" };
        foreach (var op in ops)
        {
            int idx = condition.IndexOf(op, StringComparison.Ordinal);
            if (idx > 0)
            {
                string leftPart = condition.Substring(0, idx).Trim();
                string rightPart = condition.Substring(idx + op.Length).Trim();

                // 提取字段名 [xxx]
                string fieldName = ExtractFieldName(leftPart);
                if (string.IsNullOrEmpty(fieldName)) continue;

                var cellValue = row.GetValue(fieldName);
                string cellStr = cellValue?.ToString() ?? string.Empty;

                // 提取比较值（去掉引号）
                string compareValue = rightPart.Trim('\'', '"', ' ');

                // 特殊处理空值
                if (compareValue == "")
                {
                    bool isEmpty = string.IsNullOrWhiteSpace(cellStr);
                    return op switch
                    {
                        "!=" => !isEmpty,
                        "=" => isEmpty,
                        _ => false
                    };
                }

                return op switch
                {
                    "=" => string.Equals(cellStr, compareValue, StringComparison.OrdinalIgnoreCase),
                    "!=" => !string.Equals(cellStr, compareValue, StringComparison.OrdinalIgnoreCase),
                    ">=" => CompareValues(cellValue, compareValue) >= 0,
                    "<=" => CompareValues(cellValue, compareValue) <= 0,
                    ">" => CompareValues(cellValue, compareValue) > 0,
                    "<" => CompareValues(cellValue, compareValue) < 0,
                    _ => true
                };
            }
        }

        return true;
    }

    private static string ExtractFieldName(string text)
    {
        int start = text.IndexOf('[');
        int end = text.IndexOf(']');
        if (start >= 0 && end > start)
        {
            return text.Substring(start + 1, end - start - 1);
        }
        return text;
    }

    private static int CompareValues(object? cellValue, string compareValue)
    {
        if (cellValue == null) return -1;

        if (cellValue is double d)
        {
            if (double.TryParse(compareValue, out double cmp))
                return d.CompareTo(cmp);
        }

        if (cellValue is DateTime dt)
        {
            if (DateTime.TryParse(compareValue, out DateTime cmp))
                return dt.CompareTo(cmp);
        }

        return string.Compare(cellValue.ToString(), compareValue, StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> SplitByOperator(string expr, string separator, StringComparison comparison)
    {
        var parts = new List<string>();
        int idx;
        while ((idx = expr.IndexOf(separator, comparison)) >= 0)
        {
            parts.Add(expr.Substring(0, idx));
            expr = expr.Substring(idx + separator.Length);
        }
        parts.Add(expr);
        return parts;
    }
}

/// <summary>
/// 单元格值比较器：数字按数值比、日期按时间比、其余按文本比。
/// 同一列出现混合类型时不会抛异常。
/// </summary>
public sealed class CellComparer : IComparer<object?>
{
    public static readonly CellComparer Instance = new();

    public int Compare(object? x, object? y)
    {
        if (x == null && y == null) return 0;
        if (x == null) return -1;
        if (y == null) return 1;

        if (x is DateTime dx && y is DateTime dy) return dx.CompareTo(dy);

        bool nx = ValueFormatter.TryToDouble(x, out double vx);
        bool ny = ValueFormatter.TryToDouble(y, out double vy);
        if (nx && ny) return vx.CompareTo(vy);
        if (nx) return -1;
        if (ny) return 1;

        return string.Compare(
            ValueFormatter.ToDisplayText(x),
            ValueFormatter.ToDisplayText(y),
            StringComparison.CurrentCulture);
    }
}

/// <summary>视图分组</summary>
public class ViewGroup
{
    public string Key { get; set; } = string.Empty;
    public List<DataRowModel> Rows { get; set; } = new();
}

/// <summary>视图引擎输出数据集</summary>
public class ViewDataSet
{
    public List<ViewGroup> Groups { get; set; } = new();
    public int TotalCount { get; set; }
}
