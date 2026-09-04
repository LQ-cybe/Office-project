using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MultiTableAddin.Core;

/// <summary>字段名匹配方式</summary>
public enum RuleMatchMode
{
    /// <summary>字段名包含关键词</summary>
    Contains,
    /// <summary>字段名完全等于关键词</summary>
    Equals,
    /// <summary>字段名以关键词开头</summary>
    StartsWith,
    /// <summary>字段名以关键词结尾</summary>
    EndsWith
}

public static class RuleMatchModeHelper
{
    private static readonly Dictionary<RuleMatchMode, string> Labels = new()
    {
        { RuleMatchMode.Contains, "包含" },
        { RuleMatchMode.Equals, "等于" },
        { RuleMatchMode.StartsWith, "开头是" },
        { RuleMatchMode.EndsWith, "结尾是" }
    };

    public static string GetLabel(RuleMatchMode m) => Labels.TryGetValue(m, out var l) ? l : m.ToString();

    public static IEnumerable<KeyValuePair<RuleMatchMode, string>> AllLabels => Labels;
}

/// <summary>
/// 单条字段识别规则。
/// 例如：字段名包含“部门” → 类型为单选下拉，选项为 工程部;包装部;...
/// </summary>
public class FieldRule
{
    public string Id { get; set; } = string.Empty;

    /// <summary>规则说明，用于配置表格展示</summary>
    public string Remark { get; set; } = string.Empty;

    /// <summary>匹配关键词，多个用英文分号分隔</summary>
    public string Keywords { get; set; } = string.Empty;

    public RuleMatchMode MatchMode { get; set; } = RuleMatchMode.Contains;

    public FieldType Type { get; set; } = FieldType.Text;

    /// <summary>固定选项，多个用英文分号分隔（仅单选类型有效）</summary>
    public string Options { get; set; } = string.Empty;

    /// <summary>优先级，数值越大越先匹配</summary>
    public int Priority { get; set; } = 100;

    public bool Enabled { get; set; } = true;

    /// <summary>内置规则标记；内置规则可禁用/修改但不建议删除</summary>
    public bool BuiltIn { get; set; }

    [JsonIgnore]
    public List<string> KeywordList =>
        Keywords.Split(new[] { ';', '；', ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .Where(k => k.Length > 0)
                .ToList();

    [JsonIgnore]
    public List<string> OptionList =>
        Options.Split(new[] { ';', '；', ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
               .Select(k => k.Trim())
               .Where(k => k.Length > 0)
               .ToList();

    /// <summary>判断字段名是否命中该规则</summary>
    public bool IsMatch(string fieldName)
    {
        if (!Enabled || string.IsNullOrWhiteSpace(fieldName)) return false;

        string name = fieldName.Trim().ToLowerInvariant();
        foreach (string kw in KeywordList)
        {
            string k = kw.ToLowerInvariant();
            bool hit = MatchMode switch
            {
                RuleMatchMode.Equals => name == k,
                RuleMatchMode.StartsWith => name.StartsWith(k, StringComparison.Ordinal),
                RuleMatchMode.EndsWith => name.EndsWith(k, StringComparison.Ordinal),
                _ => name.Contains(k, StringComparison.Ordinal)
            };
            if (hit) return true;
        }
        return false;
    }

    public FieldRule Clone() => (FieldRule)MemberwiseClone();
}

/// <summary>
/// 字段规则库：管理全部字段识别规则，支持用户手动增删改。
/// 存储位置 %AppData%\MultiTableAddin\field-rules.json，跨工作簿共享。
/// </summary>
public class FieldRuleLibrary
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static FieldRuleLibrary? _current;

    public string Version { get; set; } = AppVersion.ConfigSchemaVersion;

    public List<FieldRule> Rules { get; set; } = new();

    /// <summary>进程内共享的规则库单例</summary>
    public static FieldRuleLibrary Current => _current ??= Load();

    /// <summary>规则库文件路径</summary>
    public static string GetLibraryPath()
    {
        string dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MultiTableAddin");
        return Path.Combine(dir, "field-rules.json");
    }

    public static FieldRuleLibrary Load()
    {
        try
        {
            string path = GetLibraryPath();
            if (!File.Exists(path))
            {
                var def = CreateDefault();
                def.Save();
                return def;
            }

            string json = File.ReadAllText(path);
            var lib = JsonSerializer.Deserialize<FieldRuleLibrary>(json, JsonOpts);
            if (lib == null || lib.Rules.Count == 0) return CreateDefault();

            // 补齐后续版本新增的内置规则，不覆盖用户已有配置
            var defaults = CreateDefault();
            var existingIds = new HashSet<string>(lib.Rules.Select(r => r.Id));
            foreach (var r in defaults.Rules.Where(r => !existingIds.Contains(r.Id)))
                lib.Rules.Add(r);

            return lib;
        }
        catch (Exception ex)
        {
            AddInLog.Write("FieldRuleLibrary.Load.Error", ex.ToString());
            return CreateDefault();
        }
    }

    public void Save()
    {
        try
        {
            string path = GetLibraryPath();
            string? dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            Version = AppVersion.ConfigSchemaVersion;
            File.WriteAllText(path, JsonSerializer.Serialize(this, JsonOpts));
            _current = this;
            AddInLog.Write("FieldRuleLibrary.Save", $"Rules={Rules.Count} -> {path}");
        }
        catch (Exception ex)
        {
            AddInLog.Write("FieldRuleLibrary.Save.Error", ex.ToString());
        }
    }

    /// <summary>重置为内置默认规则</summary>
    public static FieldRuleLibrary ResetToDefault()
    {
        var def = CreateDefault();
        def.Save();
        return def;
    }

    /// <summary>按优先级匹配字段名，返回命中的第一条规则</summary>
    public FieldRule? Match(string fieldName)
    {
        if (string.IsNullOrWhiteSpace(fieldName)) return null;

        return Rules.Where(r => r.Enabled)
                    .OrderByDescending(r => r.Priority)
                    .ThenByDescending(r => r.KeywordList.Count == 0 ? 0 : r.KeywordList.Max(k => k.Length))
                    .FirstOrDefault(r => r.IsMatch(fieldName));
    }

    private static FieldRule R(string id, string remark, string keywords, FieldType type,
        int priority, string options = "", RuleMatchMode mode = RuleMatchMode.Contains) => new()
    {
        Id = id,
        Remark = remark,
        Keywords = keywords,
        Type = type,
        Priority = priority,
        Options = options,
        MatchMode = mode,
        Enabled = true,
        BuiltIn = true
    };

    /// <summary>内置默认规则集</summary>
    public static FieldRuleLibrary CreateDefault()
    {
        var lib = new FieldRuleLibrary();
        lib.Rules.AddRange(new[]
        {
            // ── 时间类（优先级最高，避免被“金额/数量”等误抢）─────────────
            R("rule_quarter", "季度字段固定四个选项", "季度;quarter", FieldType.Quarter, 300,
                string.Join(";", FieldSchema.QuarterOptions)),
            R("rule_datetime", "精确到时分的时间戳字段", "创建时间;更新时间;修改时间;登记时间;时间戳;提交时间", FieldType.DateTime, 290),
            R("rule_date", "常规日期字段", "日期;date;交期;生日;出生;开始日;结束日;签约日;到期", FieldType.Date, 280),

            // ── 金额 / 数值类 ──────────────────────────────────────────
            R("rule_percent", "百分比字段", "百分比;占比;比率;比例;完成率;增长率;达成率;折扣;率", FieldType.Percentage, 270),
            R("rule_currency", "金额类字段，展示为 ¥ 千分位", "金额;单价;售价;进价;成本;价格;总价;总额;收入;支出;费用;工资;薪资;报价;货款;营收;利润", FieldType.Currency, 260),
            R("rule_integer", "整数计数字段", "数量;件数;个数;台数;人数;库存;数;qty;count;张数;箱数;套数", FieldType.Integer, 250),

            // ── 联系方式 ───────────────────────────────────────────────
            R("rule_email", "邮箱字段", "邮箱;email;e-mail;邮件地址", FieldType.Email, 240),
            R("rule_phone", "电话字段", "电话;手机;联系方式;联系电话;tel;phone;mobile", FieldType.Phone, 235),
            R("rule_age", "年龄字段", "年龄;岁数;工龄", FieldType.Integer, 233),
            R("rule_url", "网址字段", "网址;链接;url;网站;主页;link", FieldType.Url, 230),

            // ── 固定选项（可由用户扩充）────────────────────────────────
            R("rule_dept", "部门字段，按企业实际部门维护选项", "部门;科室;车间;班组", FieldType.Select, 220,
                "工程部;包装部;生产部;销售部;技术部;财务部;采购部;行政部;品质部"),
            R("rule_class", "班级字段", "班级;年级;班次", FieldType.Select, 215,
                "一班;二班;三班;四班;五班;六班"),
            R("rule_status", "状态字段", "状态;进度;阶段;审批", FieldType.Select, 210,
                "未开始;进行中;已完成;已暂停;已取消"),
            R("rule_priority", "优先级字段", "优先级;紧急程度;重要度", FieldType.Select, 208,
                "高;中;低"),
            R("rule_category", "类别字段，选项由数据自动收集", "类型;类别;分类;品类;等级;级别", FieldType.Select, 205),

            // ── 文本类 ────────────────────────────────────────────────
            R("rule_longtext", "长文本字段，表单中使用多行输入框", "备注;说明;描述;详情;内容;摘要;意见;总结;方案", FieldType.LongText, 200),
            R("rule_image", "图片路径字段", "图片;照片;封面;缩略图;image;photo;头像", FieldType.Image, 195),
            R("rule_checkbox", "布尔勾选字段", "是否;已完成;启用;禁用;有效", FieldType.Checkbox, 190, string.Empty, RuleMatchMode.StartsWith),
            R("rule_name", "姓名 / 名称类字段，保持单行文本", "姓名;名称;负责人;联系人;客户;供应商;员工;操作人;经办人;name", FieldType.Text, 185),
            R("rule_id", "编号类字段，强制文本避免丢失前导零", "编号;单号;工号;学号;代码;code;编码;序号;id", FieldType.Text, 180)
        });
        return lib;
    }
}
