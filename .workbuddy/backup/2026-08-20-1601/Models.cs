using System;
using System.Collections.Generic;
using System.Linq;

namespace MultiTableAddin.Core;

/// <summary>插件版本信息（用于变更跟踪与关于对话框）</summary>
public static class AppVersion
{
    /// <summary>语义化版本号</summary>
    public const string Version = "2.9.39";

    /// <summary>版本发布日期</summary>
    public const string ReleaseDate = "2026-08-20";

    /// <summary>版本代号</summary>
    public const string CodeName = "StablePath";

    /// <summary>配置文件格式版本，格式不兼容时递增</summary>
    public const string ConfigSchemaVersion = "2.0";

    public static string DisplayText => $"v{Version} ({CodeName}) · {ReleaseDate}";
}

/// <summary>视图类型枚举</summary>
public enum ViewType
{
    Table,
    Form,
    Kanban,
    Gallery,
    Calendar,
    Gantt,
    Dashboard,
    Chart
}

/// <summary>字段类型枚举</summary>
public enum FieldType
{
    /// <summary>普通单行文本</summary>
    Text,
    /// <summary>多行长文本</summary>
    LongText,
    /// <summary>小数</summary>
    Number,
    /// <summary>整数（数量、件数等）</summary>
    Integer,
    /// <summary>日期</summary>
    Date,
    /// <summary>日期 + 时间</summary>
    DateTime,
    /// <summary>单选（下拉框）</summary>
    Select,
    /// <summary>季度</summary>
    Quarter,
    /// <summary>金额</summary>
    Currency,
    /// <summary>百分比</summary>
    Percentage,
    /// <summary>邮箱</summary>
    Email,
    /// <summary>手机号</summary>
    Phone,
    /// <summary>网址</summary>
    Url,
    /// <summary>布尔勾选</summary>
    Checkbox,
    /// <summary>图片路径</summary>
    Image
}

/// <summary>字段类型辅助方法</summary>
public static class FieldTypeHelper
{
    private static readonly Dictionary<FieldType, string> Labels = new()
    {
        { FieldType.Text, "单行文本" },
        { FieldType.LongText, "多行文本" },
        { FieldType.Number, "数字" },
        { FieldType.Integer, "整数" },
        { FieldType.Date, "日期" },
        { FieldType.DateTime, "日期时间" },
        { FieldType.Select, "单选下拉" },
        { FieldType.Quarter, "季度" },
        { FieldType.Currency, "金额" },
        { FieldType.Percentage, "百分比" },
        { FieldType.Email, "邮箱" },
        { FieldType.Phone, "手机号" },
        { FieldType.Url, "网址" },
        { FieldType.Checkbox, "勾选" },
        { FieldType.Image, "图片" }
    };

    public static string GetLabel(FieldType type) =>
        Labels.TryGetValue(type, out var label) ? label : type.ToString();

    public static IEnumerable<KeyValuePair<FieldType, string>> AllLabels => Labels;

    /// <summary>该类型是否可参与数值聚合</summary>
    public static bool IsNumeric(FieldType type) =>
        type is FieldType.Number or FieldType.Integer or FieldType.Currency or FieldType.Percentage;

    /// <summary>该类型是否为时间类型</summary>
    public static bool IsTemporal(FieldType type) =>
        type is FieldType.Date or FieldType.DateTime;

    /// <summary>该类型是否适合作为分组维度</summary>
    public static bool IsDimension(FieldType type) =>
        type is FieldType.Text or FieldType.Select or FieldType.Quarter or FieldType.Checkbox;
}

/// <summary>字段元信息</summary>
public class FieldSchema
{
    public string Name { get; set; } = string.Empty;
    public FieldType Type { get; set; } = FieldType.Text;
    public List<string> Options { get; set; } = new();

    /// <summary>Excel 列的数字格式串，用于辅助判断日期列</summary>
    public string NumberFormat { get; set; } = string.Empty;

    /// <summary>该列在 ListObject 中的序号（1 基）</summary>
    public int ColumnIndex { get; set; }

    public FieldSchema Clone() => new()
    {
        Name = Name,
        Type = Type,
        Options = new List<string>(Options),
        NumberFormat = NumberFormat,
        ColumnIndex = ColumnIndex
    };

    /// <summary>季度标准选项</summary>
    public static readonly List<string> QuarterOptions = new()
    {
        "第一季度", "第二季度", "第三季度", "第四季度"
    };

    /// <summary>判断字符串是否为季度值</summary>
    public static bool IsQuarterValue(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;
        string v = s.Trim().ToLowerInvariant();
        return v is "第一季度" or "第二季度" or "第三季度" or "第四季度"
            or "一季度" or "二季度" or "三季度" or "四季度"
            or "q1" or "q2" or "q3" or "q4"
            or "1季度" or "2季度" or "3季度" or "4季度"
            or "第1季度" or "第2季度" or "第3季度" or "第4季度";
    }

    /// <summary>把任意季度写法规范化为标准季度选项</summary>
    public static string NormalizeQuarter(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        string v = s.Trim().ToLowerInvariant();
        if (v.Contains('1') || v.Contains('一')) return "第一季度";
        if (v.Contains('2') || v.Contains('二')) return "第二季度";
        if (v.Contains('3') || v.Contains('三')) return "第三季度";
        if (v.Contains('4') || v.Contains('四')) return "第四季度";
        return s.Trim();
    }
}

/// <summary>单行数据记录</summary>
public class DataRowModel
{
    /// <summary>在 ListObject DataBodyRange 中的行号（1 基）</summary>
    public int RowIndex { get; set; }

    /// <summary>字段原始值（用于编辑、排序、计算）</summary>
    public Dictionary<string, object?> Values { get; set; } = new();

    /// <summary>Excel 实际显示文本（用于表格视图按 Excel 所见即所得渲染）</summary>
    public Dictionary<string, string> DisplayTexts { get; set; } = new();

    public object? GetValue(string fieldName) =>
        fieldName != null && Values.TryGetValue(fieldName, out var v) ? v : null;

    public void SetValue(string fieldName, object? value) =>
        Values[fieldName] = value;

    public string GetText(string fieldName) =>
        DisplayTexts.TryGetValue(fieldName, out var text)
            ? text
            : ValueFormatter.ToDisplayText(GetValue(fieldName));

    /// <summary>获取 Excel 显示文本；不存在时按字段类型格式化</summary>
    public string GetDisplayText(string fieldName, FieldSchema? field = null)
    {
        if (DisplayTexts.TryGetValue(fieldName, out var text)) return text;
        var value = GetValue(fieldName);
        if (field != null) return ValueFormatter.ToDisplayText(value, field.Type);
        return ValueFormatter.ToDisplayText(value);
    }
}

/// <summary>内存数据表</summary>
public class DataTableModel
{
    public string SourceFile { get; set; } = string.Empty;
    public string SheetName { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;
    public List<FieldSchema> Fields { get; set; } = new();
    public List<DataRowModel> Rows { get; set; } = new();
    public bool IsDirty { get; set; }

    public DataRowModel? FindRow(int rowIndex) => Rows.Find(r => r.RowIndex == rowIndex);

    public FieldSchema? FindField(string name) => Fields.Find(f => f.Name == name);

    public List<string> FieldNames => Fields.ConvertAll(f => f.Name);
}

/// <summary>数据源描述（工作表 + 超级表）</summary>
public class TableSourceInfo
{
    public string SheetName { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }

    public string DisplayText => $"{SheetName} / {TableName}  ({RowCount} 行 × {ColumnCount} 列)";

    public string Key => SheetName + "\u0001" + TableName;
}

/// <summary>排序配置</summary>
public class SortConfig
{
    public string Field { get; set; } = string.Empty;
    public string Order { get; set; } = "asc"; // asc / desc
}

/// <summary>字段级筛选配置（单字段关键词筛选）</summary>
public class FieldFilterConfig
{
    public string Field { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

/// <summary>卡片布局元信息</summary>
public class CardMeta
{
    public string Title { get; set; } = string.Empty;
    public string Image { get; set; } = string.Empty;
    public List<string> Description { get; set; } = new();

    /// <summary>按该字段值自动着色（卡片视图分色显示）</summary>
    public string ColorField { get; set; } = string.Empty;

    /// <summary>图片文件夹模式：指定文件夹路径，按匹配字段值查找同名图片文件</summary>
    public string ImageFolder { get; set; } = string.Empty;

    /// <summary>图片文件夹模式下，用该字段的值去匹配图片文件名（不含扩展名）</summary>
    public string ImageMatchField { get; set; } = string.Empty;

    /// <summary>图片显示模式：normal=正常矩形；circle=圆形居中裁剪（适合人员照片类）</summary>
    public string ImageShape { get; set; } = "normal";
}

/// <summary>甘特图配置</summary>
public class GanttConfig
{
    public string StartField { get; set; } = string.Empty;
    public string EndField { get; set; } = string.Empty;
    public string LabelField { get; set; } = string.Empty;

    /// <summary>分组字段（左侧分组树）</summary>
    public string GroupField { get; set; } = string.Empty;

    /// <summary>进度字段（0~1 或 0~100）</summary>
    public string ProgressField { get; set; } = string.Empty;

    /// <summary>甘特条颜色字段（按该字段值自动着色）</summary>
    public string ColorField { get; set; } = string.Empty;

    /// <summary>时间维度：年/季度/月/周/日</summary>
    public TimeDimension TimeDimension { get; set; } = TimeDimension.Day;

    /// <summary>卡片/提示中额外显示的字段</summary>
    public List<string> DisplayFields { get; set; } = new();
}

/// <summary>日历配置</summary>
public class CalendarConfig
{
    public string DateField { get; set; } = string.Empty;
    public string EndDateField { get; set; } = string.Empty;
    public string TitleField { get; set; } = string.Empty;

    /// <summary>按该字段值自动着色（日历分色显示）</summary>
    public string ColorField { get; set; } = string.Empty;

    /// <summary>拖拽调整时间时的吸附步长（分钟）</summary>
    public int SnapMinutes { get; set; } = 15;
}

/// <summary>表格视图配置</summary>
public class TableViewConfig
{
    /// <summary>是否显示序号列</summary>
    public bool ShowRowNumber { get; set; } = true;

    /// <summary>列宽自适应下限（像素）</summary>
    public double MinColumnWidth { get; set; } = 60;

    /// <summary>列宽自适应上限（像素），避免长文本撑爆表格</summary>
    public double MaxColumnWidth { get; set; } = 320;

    /// <summary>参与宽度测量的最大采样行数</summary>
    public int WidthSampleRows { get; set; } = 200;
}

/// <summary>图表类型枚举</summary>
public enum ChartType
{
    /// <summary>纵向柱状图</summary>
    Column,
    /// <summary>横向条形图</summary>
    Bar,
    /// <summary>折线图</summary>
    Line,
    /// <summary>面积图</summary>
    Area,
    /// <summary>饼图</summary>
    Pie,
    /// <summary>环形图</summary>
    Doughnut,
    /// <summary>仪表盘</summary>
    Gauge
}

/// <summary>时间维度枚举</summary>
public enum TimeDimension
{
    None,
    Year,
    Quarter,
    Month,
    Week,
    Day
}

/// <summary>表格汇总行位置</summary>
public enum SummaryPosition
{
    /// <summary>固定在数据末尾（默认）</summary>
    DataEnd,
    /// <summary>吸附在窗口底部</summary>
    WindowEnd
}

/// <summary>聚合方式</summary>
public enum AggregateMode
{
    Sum,
    Count,
    Average,
    Max,
    Min,
    DistinctCount
}

public static class AggregateModeHelper
{
    private static readonly Dictionary<AggregateMode, string> Labels = new()
    {
        { AggregateMode.Sum, "求和" },
        { AggregateMode.Count, "计数" },
        { AggregateMode.Average, "平均值" },
        { AggregateMode.Max, "最大值" },
        { AggregateMode.Min, "最小值" },
        { AggregateMode.DistinctCount, "去重计数" }
    };

    public static string GetLabel(AggregateMode mode) =>
        Labels.TryGetValue(mode, out var l) ? l : mode.ToString();

    public static IEnumerable<KeyValuePair<AggregateMode, string>> AllLabels => Labels;
}

public static class TimeDimensionHelper
{
    private static readonly Dictionary<TimeDimension, string> Labels = new()
    {
        { TimeDimension.None, "不按时间" },
        { TimeDimension.Year, "按年" },
        { TimeDimension.Quarter, "按季度" },
        { TimeDimension.Month, "按月" },
        { TimeDimension.Week, "按周" },
        { TimeDimension.Day, "按日" }
    };

    public static string GetLabel(TimeDimension d) =>
        Labels.TryGetValue(d, out var l) ? l : d.ToString();

    public static IEnumerable<KeyValuePair<TimeDimension, string>> AllLabels => Labels;
}

public static class ChartTypeHelper
{
    private static readonly Dictionary<ChartType, string> Labels = new()
    {
        { ChartType.Column, "柱状图" },
        { ChartType.Bar, "条形图" },
        { ChartType.Line, "折线图" },
        { ChartType.Area, "面积图" },
        { ChartType.Pie, "饼图" },
        { ChartType.Doughnut, "环形图" },
        { ChartType.Gauge, "仪表盘" }
    };

    public static string GetLabel(ChartType t) =>
        Labels.TryGetValue(t, out var l) ? l : t.ToString();

    public static IEnumerable<KeyValuePair<ChartType, string>> AllLabels => Labels;
}

/// <summary>单个图表的配置</summary>
public class ChartConfig
{
    public string Id { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public ChartType Type { get; set; } = ChartType.Column;

    /// <summary>分组维度字段（非时间维度时使用）</summary>
    public string DimensionField { get; set; } = string.Empty;

    /// <summary>度量字段</summary>
    public string MetricField { get; set; } = string.Empty;

    public AggregateMode Aggregation { get; set; } = AggregateMode.Sum;

    /// <summary>时间维度字段；设置后按 TimeGroup 聚合</summary>
    public string TimeField { get; set; } = string.Empty;

    public TimeDimension TimeGroup { get; set; } = TimeDimension.None;

    /// <summary>可选的次级分组字段，形成多系列</summary>
    public string SeriesField { get; set; } = string.Empty;

    /// <summary>最多展示的分类数量，超出合并为“其他”</summary>
    public int TopN { get; set; } = 12;

    /// <summary>仪表盘目标值</summary>
    public double GaugeTarget { get; set; } = 100;

    /// <summary>在仪表盘布局中占据的列数（1 或 2）</summary>
    public int ColumnSpan { get; set; } = 1;

    /// <summary>图表区域高度</summary>
    public double Height { get; set; } = 260;

    /// <summary>小于等于该值的数据标签将被隐藏（0 表示隐藏 0 及负数）</summary>
    public double MinLabelValue { get; set; } = 0;

    public ChartConfig Clone() => (ChartConfig)MemberwiseClone();
}

/// <summary>仪表盘 KPI 卡片配置</summary>
public class StatCardConfig
{
    public string Id { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Field { get; set; } = string.Empty;
    public AggregateMode Aggregation { get; set; } = AggregateMode.Count;

    /// <summary>可选：只统计满足该筛选表达式的行</summary>
    public string Filter { get; set; } = string.Empty;

    /// <summary>显示格式：auto / int / money / percent</summary>
    public string Format { get; set; } = "auto";

    /// <summary>卡片强调色（#RRGGBB），为空则按序取主题色</summary>
    public string? Color { get; set; }
}

/// <summary>仪表盘配置</summary>
public class DashboardConfig
{
    public List<StatCardConfig> StatCards { get; set; } = new();
    public List<ChartConfig> Charts { get; set; } = new();

    /// <summary>图表区列数</summary>
    public int Columns { get; set; } = 2;
}

/// <summary>字段手动配置覆盖（允许用户修改字段类型、选项及校验规则）</summary>
public class FieldOverride
{
    public string Name { get; set; } = string.Empty;
    public FieldType Type { get; set; } = FieldType.Text;
    public List<string> Options { get; set; } = new();

    /// <summary>是否为用户显式设定，true 时不再被自动推断覆盖</summary>
    public bool UserDefined { get; set; } = true;

    // ── 显示格式 ──
    public string Format { get; set; } = string.Empty;

    // ── 校验规则 ──
    public bool Required { get; set; }
    public double? MinValue { get; set; }
    public double? MaxValue { get; set; }

    /// <summary>数值/整数/金额/百分比等数字字段的步进值（上下箭头增量）</summary>
    public double? Step { get; set; }

    public int? MinLength { get; set; }
    public int? MaxLength { get; set; }
    public string RegexPattern { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
}

/// <summary>数量 / 单价 / 金额 三向联动配置</summary>
public class NumericLinkConfig
{
    public string QuantityField { get; set; } = string.Empty;
    public string UnitPriceField { get; set; } = string.Empty;
    public string AmountField { get; set; } = string.Empty;

    /// <summary>金额保留小数位</summary>
    public int AmountDecimals { get; set; } = 2;

    /// <summary>单价保留小数位</summary>
    public int UnitPriceDecimals { get; set; } = 4;

    public bool IsValid =>
        !string.IsNullOrWhiteSpace(QuantityField) &&
        !string.IsNullOrWhiteSpace(UnitPriceField) &&
        !string.IsNullOrWhiteSpace(AmountField);

    /// <summary>判断字段是否参与联动</summary>
    public bool Involves(string fieldName) =>
        fieldName == QuantityField || fieldName == UnitPriceField || fieldName == AmountField;
}

/// <summary>单个视图配置</summary>
public class ViewConfig
{
    public string ViewId { get; set; } = string.Empty;
    public ViewType ViewType { get; set; } = ViewType.Table;
    public string ViewName { get; set; } = string.Empty;
    public string Filter { get; set; } = string.Empty;
    public FieldFilterConfig? FieldFilter { get; set; }
    public string GroupBy { get; set; } = string.Empty;
    public List<SortConfig> Sort { get; set; } = new();
    public List<string> VisibleFields { get; set; } = new();
    public CardMeta? CardMeta { get; set; }
    public GanttConfig? GanttConfig { get; set; }
    public CalendarConfig? CalendarConfig { get; set; }
    public ChartConfig? ChartConfig { get; set; }
    public DashboardConfig? DashboardConfig { get; set; }
    public TableViewConfig? TableConfig { get; set; }

    /// <summary>表格视图底部汇总行：字段名 → 计算方式（对应 JS 的 v.summary）</summary>
    public Dictionary<string, string> Summary { get; set; } = new();

    /// <summary>是否显示底部汇总行（对应 JS 的 v.showSummary）</summary>
    public bool ShowSummary { get; set; } = false;

    /// <summary>汇总行位置：数据末尾（默认）或窗口底部</summary>
    public SummaryPosition SummaryPosition { get; set; } = SummaryPosition.DataEnd;

    /// <summary>是否为用户创建的视图（由「另存新视图」或动态补建产生）。
    /// 自带基础视图固定为 false，不可删除、不可重命名；用户视图可删、可改名。
    /// 旧配置缺该字段时反序列化为 false，天然等价于基础视图，向后兼容。</summary>
    public bool UserView { get; set; } = false;

    public static string NewId(string prefix) =>
        prefix + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
}

/// <summary>视图配置文件根模型</summary>
public class ViewConfigFile
{
    public string Version { get; set; } = AppVersion.ConfigSchemaVersion;
    public string GeneratedBy { get; set; } = "MultiTableAddin " + AppVersion.Version;
    public string SourceFile { get; set; } = string.Empty;
    public string SourceSheet { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;
    public List<FieldSchema> Fields { get; set; } = new();
    public List<FieldOverride> FieldOverrides { get; set; } = new();
    public List<NumericLinkConfig> NumericLinks { get; set; } = new();
    public List<ViewConfig> Views { get; set; } = new();

    /// <summary>获取字段的有效类型（优先使用用户覆盖配置）</summary>
    public FieldType GetEffectiveFieldType(string fieldName) => GetEffectiveField(fieldName).Type;

    /// <summary>获取字段的有效选项（优先使用用户覆盖配置）</summary>
    public List<string> GetEffectiveOptions(string fieldName) => GetEffectiveField(fieldName).Options;

    /// <summary>获取字段的有效 FieldSchema（合并自动推断和用户覆盖）</summary>
    public FieldSchema GetEffectiveField(string fieldName)
    {
        var field = Fields.Find(f => f.Name == fieldName);
        var ov = FieldOverrides?.Find(f => f.Name == fieldName);

        if (field == null)
        {
            return ov == null
                ? new FieldSchema { Name = fieldName }
                : new FieldSchema { Name = fieldName, Type = ov.Type, Options = new List<string>(ov.Options) };
        }

        if (ov == null) return field;

        return new FieldSchema
        {
            Name = field.Name,
            Type = ov.Type,
            Options = ov.Options.Count > 0 ? new List<string>(ov.Options) : new List<string>(field.Options),
            NumberFormat = field.NumberFormat,
            ColumnIndex = field.ColumnIndex
        };
    }

    /// <summary>设置字段覆盖配置</summary>
    public void SetFieldOverride(string fieldName, FieldType type, List<string> options)
    {
        FieldOverrides ??= new List<FieldOverride>();
        var ov = FieldOverrides.Find(f => f.Name == fieldName);
        if (ov == null)
        {
            ov = new FieldOverride { Name = fieldName };
            FieldOverrides.Add(ov);
        }
        ov.Type = type;
        ov.Options = options ?? new List<string>();
        ov.UserDefined = true;
    }

    public void RemoveFieldOverride(string fieldName) =>
        FieldOverrides?.RemoveAll(f => f.Name == fieldName);

    /// <summary>查找包含指定字段的数值联动配置</summary>
    public NumericLinkConfig? FindNumericLink(string fieldName) =>
        NumericLinks?.FirstOrDefault(l => l.IsValid && l.Involves(fieldName));

    /// <summary>获取字段的覆盖配置（含校验规则），不存在返回 null</summary>
    public FieldOverride? GetFieldOverride(string fieldName) =>
        FieldOverrides?.Find(f => f.Name == fieldName);
}

/// <summary>
/// #428-3 工作簿级合并配置：一个 Excel 文件对应唯一一个 .multiview.json，
/// 内含多张超级表（ListObject）各自的 ViewConfigFile，以及上次打开的表名（记忆恢复用）。
/// 该结构同时被序列化进 Excel 隐藏工作表 _MultiTableConfig（#428-4）。
/// </summary>
public class WorkbookConfigFile
{
    public string Version { get; set; } = AppVersion.ConfigSchemaVersion;
    public string GeneratedBy { get; set; } = "MultiTableAddin " + AppVersion.Version;
    public string SourceFile { get; set; } = string.Empty;
    /// <summary>上次使用的超级表名；重新打开工作簿时自动恢复到此表</summary>
    public string LastTableName { get; set; } = string.Empty;
    /// <summary>超级表名 -> 该表的视图配置</summary>
    public Dictionary<string, ViewConfigFile> Tables { get; set; } = new();
    /// <summary>超级表展示顺序（列表中的先后顺序）</summary>
    public List<string> Order { get; set; } = new();
}
