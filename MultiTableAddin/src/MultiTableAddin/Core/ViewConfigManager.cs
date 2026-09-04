using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MultiTableAddin.Core;

/// <summary>
/// 视图配置管理器：加载/保存 .multiview.json 配置文件
/// 配置文件与 xlsx 同目录，命名 {文件名}.multiview.json
/// </summary>
public class ViewConfigManager
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        Converters = { new JsonStringEnumConverter() },
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    /// <summary>
    /// #428-3 合并配置文件路径：一个工作簿对应唯一一个 {文件名}.multiview.json（不再按表拆分）。
    /// tableName 参数保留仅为兼容旧调用，已不再参与文件名。
    /// </summary>
    public static string GetConfigFilePath(string workbookPath, string? tableName = null)
    {
        if (string.IsNullOrEmpty(workbookPath))
            return string.Empty;

        string dir = Path.GetDirectoryName(workbookPath) ?? string.Empty;
        string baseName = Path.GetFileNameWithoutExtension(workbookPath);
        return Path.Combine(dir, baseName + ".multiview.json");
    }

    private static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "table";
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new System.Text.StringBuilder(name.Length);
        foreach (char c in name)
        {
            sb.Append(invalid.Contains(c) ? '_' : c);
        }
        string result = sb.ToString().Trim();
        return string.IsNullOrEmpty(result) ? "table" : result;
    }

    /// <summary>
    /// #428-3 加载指定超级表的视图配置。
    /// 读取顺序：① Excel 隐藏工作表（#428-4）→ ② 外部合并 JSON 文件（含旧分表迁移）。
    /// 返回空配置时调用方负责生成默认配置。
    /// </summary>
    public ViewConfigFile Load(string workbookPath, string? tableName = null)
    {
        try
        {
            var wb = LoadWorkbook(workbookPath);
            if (wb != null && !string.IsNullOrWhiteSpace(tableName) && wb.Tables.TryGetValue(tableName, out var tbl))
            {
                tbl.SourceFile = Path.GetFileName(workbookPath);
                tbl.TableName = tableName;
                EnsureViewIds(tbl);
                AddInLog.Write("ViewConfigManager.Load", $"Loaded {tbl.Views.Count} views for table '{tableName}'");
                return tbl;
            }

            AddInLog.Write("ViewConfigManager.Load", "配置不存在，返回空配置: " + workbookPath + " / " + tableName);
            return new ViewConfigFile { SourceFile = Path.GetFileName(workbookPath), TableName = tableName ?? string.Empty };
        }
        catch (Exception ex)
        {
            AddInLog.Write("ViewConfigManager.Load.Error", ex.ToString());
            return new ViewConfigFile { SourceFile = Path.GetFileName(workbookPath) };
        }
    }

    /// <summary>
    /// #428-3/#428-4 保存指定超级表的视图配置（合并进单个工作簿级配置）。
    /// 根据 saveLocation 决定写入目标：
    ///   "excel" → 仅写入 Excel 隐藏工作表 _MultiTableConfig（并删除外部 JSON）
    ///   "file"  → 仅写入外部合并 JSON 文件（并删除嵌入的隐藏工作表，避免 LoadWorkbook 隐藏表优先级读到陈旧配置）
    ///   "both"  → 同时写入外部文件 + 隐藏工作表（默认）
    /// 并记忆 LastTableName（下次打开自动恢复）。
    /// </summary>
    public void Save(string workbookPath, ViewConfigFile config, string? tableName = null, string saveLocation = "both")
    {
        try
        {
            string key = tableName ?? config.TableName ?? string.Empty;
            if (string.IsNullOrWhiteSpace(key)) return;

            var wb = LoadWorkbook(workbookPath) ?? new WorkbookConfigFile { SourceFile = Path.GetFileName(workbookPath) };
            wb.Tables ??= new Dictionary<string, ViewConfigFile>();
            wb.Order ??= new List<string>();

            config.SourceFile = Path.GetFileName(workbookPath);
            config.TableName = key;
            config.Version = AppVersion.ConfigSchemaVersion;
            config.GeneratedBy = "MultiTableAddin " + AppVersion.Version;

            wb.Tables[key] = config;
            if (!wb.Order.Contains(key)) wb.Order.Add(key);
            wb.LastTableName = key;     // #428-3 记忆最后使用的超级表
            wb.Version = AppVersion.ConfigSchemaVersion;
            wb.GeneratedBy = "MultiTableAddin " + AppVersion.Version;
            wb.SourceFile = Path.GetFileName(workbookPath);

            switch ((saveLocation ?? "both").ToLowerInvariant())
            {
                case "excel":
                    SaveWorkbookToSheet(wb);
                    DeleteExternalConfigFile(workbookPath);
                    break;
                case "file":
                    SaveWorkbookToFile(workbookPath, wb);
                    DeleteConfigSheet();
                    break;
                default:
                    SaveWorkbookToFile(workbookPath, wb);
                    SaveWorkbookToSheet(wb);
                    break;
            }

            AddInLog.Write("ViewConfigManager.Save", $"Saved table '{key}' (tables={wb.Tables.Count}, location={saveLocation ?? "both"})");
        }
        catch (Exception ex)
        {
            AddInLog.Write("ViewConfigManager.Save.Error", ex.ToString());
        }
    }

    /// <summary>#462 配置保存位置设为「外部文件」时，删除嵌入的隐藏工作表</summary>
    private void DeleteConfigSheet()
    {
        try { new ExcelAdapter().DeleteConfigSheet(); }
        catch (Exception ex) { AddInLog.Write("ViewConfigManager.DeleteConfigSheet.Error", ex.ToString()); }
    }

    /// <summary>#462 配置保存位置设为「嵌入 Excel」时，删除外部 JSON 文件（单一数据源）</summary>
    private void DeleteExternalConfigFile(string workbookPath)
    {
        try
        {
            string path = GetConfigFilePath(workbookPath);
            if (!string.IsNullOrEmpty(path) && File.Exists(path)) File.Delete(path);
        }
        catch (Exception ex) { AddInLog.Write("ViewConfigManager.DeleteExternalConfigFile.Error", ex.ToString()); }
    }

    /// <summary>#428-3 读取上次使用的超级表名（供启动时自动恢复）</summary>
    public static string GetLastTableName(string workbookPath)
    {
        try
        {
            var wb = new ViewConfigManager().LoadWorkbook(workbookPath);
            return wb?.LastTableName ?? string.Empty;
        }
        catch { return string.Empty; }
    }

    // ── 工作簿级合并配置读写 ──────────────────────────────────

    /// <summary>统一加载工作簿级配置：优先 Excel 隐藏工作表，其次外部 JSON 文件</summary>
    private WorkbookConfigFile? LoadWorkbook(string workbookPath)
    {
        // ① 隐藏工作表（若工作簿已打开）
        try
        {
            var json = new ExcelAdapter().ReadConfigSheet();
            if (!string.IsNullOrEmpty(json))
            {
                var w = JsonSerializer.Deserialize<WorkbookConfigFile>(json, JsonOpts);
                if (w != null) { w.Tables ??= new Dictionary<string, ViewConfigFile>(); w.Order ??= new List<string>(); return w; }
            }
        }
        catch (Exception ex) { AddInLog.Write("ViewConfigManager.LoadWorkbook.Embed", ex.ToString()); }

        // ② 外部 JSON 文件（含旧分表迁移）
        return LoadWorkbookFromFile(workbookPath);
    }

    private WorkbookConfigFile? LoadWorkbookFromFile(string workbookPath)
    {
        string path = GetConfigFilePath(workbookPath);
        if (!File.Exists(path))
        {
            // #428-3 兼容旧版：把 {文件名}-{表名}.multiview.json 合并成单文件
            var migrated = MigrateLegacyFiles(workbookPath);
            if (migrated != null) { SaveWorkbookToFile(workbookPath, migrated); return migrated; }
            return null;
        }
        try
        {
            var w = JsonSerializer.Deserialize<WorkbookConfigFile>(File.ReadAllText(path), JsonOpts);
            if (w != null) { w.Tables ??= new Dictionary<string, ViewConfigFile>(); w.Order ??= new List<string>(); }
            return w;
        }
        catch (Exception ex) { AddInLog.Write("ViewConfigManager.LoadWorkbookFromFile.Error", ex.ToString()); return null; }
    }

    /// <summary>#428-3 归并旧版「每表一文件」到合并结构</summary>
    private WorkbookConfigFile? MigrateLegacyFiles(string workbookPath)
    {
        try
        {
            string dir = Path.GetDirectoryName(workbookPath) ?? string.Empty;
            string baseName = Path.GetFileNameWithoutExtension(workbookPath);
            var files = Directory.GetFiles(dir, baseName + "-*.multiview.json");
            if (files.Length == 0) return null;

            var wb = new WorkbookConfigFile { SourceFile = Path.GetFileName(workbookPath) };
            wb.Tables = new Dictionary<string, ViewConfigFile>();
            wb.Order = new List<string>();
            foreach (var f in files)
            {
                try
                {
                    var tbl = JsonSerializer.Deserialize<ViewConfigFile>(File.ReadAllText(f), JsonOpts);
                    if (tbl == null) continue;
                    string key = tbl.TableName;
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        var nm = Path.GetFileNameWithoutExtension(f);
                        if (nm.Length > baseName.Length + 1) key = nm.Substring(baseName.Length + 1);
                    }
                    if (string.IsNullOrWhiteSpace(key)) continue;
                    wb.Tables[key] = tbl;
                    if (!wb.Order.Contains(key)) wb.Order.Add(key);
                }
                catch { }
            }
            return wb.Tables.Count > 0 ? wb : null;
        }
        catch { return null; }
    }

    private void SaveWorkbookToFile(string workbookPath, WorkbookConfigFile wb)
    {
        string path = GetConfigFilePath(workbookPath);
        if (string.IsNullOrEmpty(path)) return;
        try
        {
            var json = JsonSerializer.Serialize(wb, JsonOpts);
            string tmp = path + ".tmp";
            File.WriteAllText(tmp, json);
            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }
        catch (Exception ex) { AddInLog.Write("ViewConfigManager.SaveWorkbookToFile.Error", ex.ToString()); }
    }

    private void SaveWorkbookToSheet(WorkbookConfigFile wb)
    {
        try
        {
            var json = JsonSerializer.Serialize(wb, JsonOpts);
            new ExcelAdapter().WriteConfigSheet(json);
        }
        catch (Exception ex) { AddInLog.Write("ViewConfigManager.SaveWorkbookToSheet.Error", ex.ToString()); }
    }

    private static void EnsureViewIds(ViewConfigFile config)
    {
        config.Views ??= new List<ViewConfig>();
        config.Fields ??= new List<FieldSchema>();
        config.FieldOverrides ??= new List<FieldOverride>();
        config.NumericLinks ??= new List<NumericLinkConfig>();
        foreach (var v in config.Views)
        {
            if (string.IsNullOrWhiteSpace(v.ViewId))
                v.ViewId = ViewConfig.NewId(v.ViewType.ToString().ToLowerInvariant());
            if (string.IsNullOrWhiteSpace(v.ViewName))
                v.ViewName = DefaultViewName(v.ViewType);
        }
    }

    /// <summary>
    /// 把最新读到的表结构同步进配置，保留用户已有设置。
    /// 修复：旧实现直接 config.Fields = table.Fields 且不更新 VisibleFields，
    /// 导致「表重建后只显示 4 个字段」。
    /// </summary>
    public static void SyncFields(ViewConfigFile config, DataTableModel table)
    {
        if (config == null || table == null) return;

        config.Fields ??= new List<FieldSchema>();
        config.FieldOverrides ??= new List<FieldOverride>();
        config.Views ??= new List<ViewConfig>();

        var currentNames = table.Fields.Select(f => f.Name).ToList();
        var currentSet = new HashSet<string>(currentNames, StringComparer.Ordinal);
        var previousSet = new HashSet<string>(config.Fields.Select(f => f.Name), StringComparer.Ordinal);

        // 1) 字段清单以实际表结构为准（推断类型来自 Excel，用户覆盖单独存放不受影响）
        config.Fields = table.Fields.Select(f => f.Clone()).ToList();

        // 2) 清理已不存在字段的覆盖配置
        config.FieldOverrides.RemoveAll(o => !currentSet.Contains(o.Name));

        // 3) 清理失效的数值联动
        config.NumericLinks ??= new List<NumericLinkConfig>();
        config.NumericLinks.RemoveAll(l =>
            !currentSet.Contains(l.QuantityField) ||
            !currentSet.Contains(l.UnitPriceField) ||
            !currentSet.Contains(l.AmountField));

        // 4) 逐视图同步可见字段：删掉消失的，追加新增的
        var newFields = currentNames.Where(n => !previousSet.Contains(n)).ToList();
        foreach (var view in config.Views)
        {
            view.VisibleFields ??= new List<string>();

            if (view.VisibleFields.Count == 0)
            {
                view.VisibleFields = new List<string>(currentNames);
            }
            else
            {
                view.VisibleFields.RemoveAll(n => !currentSet.Contains(n));
                foreach (var n in newFields)
                    if (!view.VisibleFields.Contains(n))
                        view.VisibleFields.Add(n);

                // 极端情况：全被清空则回落到全字段，避免出现空白视图
                if (view.VisibleFields.Count == 0)
                    view.VisibleFields = new List<string>(currentNames);
            }

            // 排序/分组字段失效则移除
            view.Sort?.RemoveAll(s => !currentSet.Contains(s.Field));
            if (!string.IsNullOrEmpty(view.GroupBy) && !currentSet.Contains(view.GroupBy))
                view.GroupBy = string.Empty;

            SyncViewMeta(view, table, currentSet);
        }

        AddInLog.Write("ViewConfigManager.SyncFields",
            $"Fields={currentNames.Count}, New={newFields.Count}, Views={config.Views.Count}");
    }

    private static void SyncViewMeta(ViewConfig view, DataTableModel table, HashSet<string> currentSet)
    {
        if (view.CardMeta != null)
        {
            if (!currentSet.Contains(view.CardMeta.Title))
                view.CardMeta.Title = table.Fields.Count > 0 ? table.Fields[0].Name : string.Empty;
            if (!string.IsNullOrEmpty(view.CardMeta.Image) && !currentSet.Contains(view.CardMeta.Image))
                view.CardMeta.Image = string.Empty;
            view.CardMeta.Description?.RemoveAll(d => !currentSet.Contains(d));
        }

        if (view.CalendarConfig != null && !currentSet.Contains(view.CalendarConfig.DateField))
        {
            var d = table.Fields.Find(f => FieldTypeHelper.IsTemporal(f.Type));
            view.CalendarConfig.DateField = d?.Name ?? string.Empty;
        }

        if (view.GanttConfig != null)
        {
            if (!currentSet.Contains(view.GanttConfig.StartField)) view.GanttConfig.StartField = string.Empty;
            if (!currentSet.Contains(view.GanttConfig.EndField)) view.GanttConfig.EndField = string.Empty;
            if (!currentSet.Contains(view.GanttConfig.LabelField))
                view.GanttConfig.LabelField = table.Fields.Count > 0 ? table.Fields[0].Name : string.Empty;
        }

        if (view.ChartConfig != null) SyncChart(view.ChartConfig, currentSet);

        if (view.DashboardConfig != null)
        {
            view.DashboardConfig.StatCards?.RemoveAll(c =>
                !string.IsNullOrEmpty(c.Field) && !currentSet.Contains(c.Field));
            if (view.DashboardConfig.Charts != null)
                foreach (var ch in view.DashboardConfig.Charts) SyncChart(ch, currentSet);
        }
    }

    private static void SyncChart(ChartConfig chart, HashSet<string> currentSet)
    {
        if (!string.IsNullOrEmpty(chart.DimensionField) && !currentSet.Contains(chart.DimensionField))
            chart.DimensionField = string.Empty;
        if (!string.IsNullOrEmpty(chart.MetricField) && !currentSet.Contains(chart.MetricField))
            chart.MetricField = string.Empty;
        if (!string.IsNullOrEmpty(chart.TimeField) && !currentSet.Contains(chart.TimeField))
        {
            chart.TimeField = string.Empty;
            chart.TimeGroup = TimeDimension.None;
        }
        if (!string.IsNullOrEmpty(chart.SeriesField) && !currentSet.Contains(chart.SeriesField))
            chart.SeriesField = string.Empty;
    }

    public static string DefaultViewName(ViewType type) => type switch
    {
        ViewType.Table => "表格视图",
        ViewType.Form => "表单视图",
        ViewType.Kanban => "看板视图",
        ViewType.Gallery => "画册视图",
        ViewType.Calendar => "日历视图",
        ViewType.Gantt => "甘特视图",
        ViewType.Dashboard => "仪表盘",
        ViewType.Chart => "统计图表",
        _ => "新视图"
    };

    // ─────────────────────────────────────────────────────────────
    // 默认配置生成
    // ─────────────────────────────────────────────────────────────

    /// <summary>创建默认视图配置（首次打开时自动生成）</summary>
    public ViewConfigFile CreateDefaultConfig(DataTableModel table)
    {
        var config = new ViewConfigFile
        {
            Version = AppVersion.ConfigSchemaVersion,
            GeneratedBy = "MultiTableAddin " + AppVersion.Version,
            SourceFile = Path.GetFileName(table.SourceFile),
            SourceSheet = table.SheetName,
            TableName = table.TableName,
            Fields = table.Fields.Select(f => f.Clone()).ToList()
        };

        var allNames = table.FieldNames;
        var firstName = allNames.Count > 0 ? allNames[0] : string.Empty;

        // ── 表格视图 ────────────────────────────────────────────
        config.Views.Add(new ViewConfig
        {
            ViewId = ViewConfig.NewId("table"),
            ViewType = ViewType.Table,
            ViewName = "表格视图",
            VisibleFields = new List<string>(allNames),
            Sort = new List<SortConfig>(),
            TableConfig = new TableViewConfig()
        });

        // ── 表单视图 ────────────────────────────────────────────
        config.Views.Add(new ViewConfig
        {
            ViewId = ViewConfig.NewId("form"),
            ViewType = ViewType.Form,
            ViewName = "表单视图",
            VisibleFields = new List<string>(allNames)
        });

        // ── 画册视图 ────────────────────────────────────────────
        config.Views.Add(new ViewConfig
        {
            ViewId = ViewConfig.NewId("gallery"),
            ViewType = ViewType.Gallery,
            ViewName = "画册视图",
            VisibleFields = new List<string>(allNames),
            CardMeta = BuildCardMeta(table)
        });

        // ── 日历视图（需要日期字段）────────────────────────────
        var dateFields = table.Fields.FindAll(f => FieldTypeHelper.IsTemporal(f.Type));
        if (dateFields.Count > 0)
        {
            config.Views.Add(new ViewConfig
            {
                ViewId = ViewConfig.NewId("calendar"),
                ViewType = ViewType.Calendar,
                ViewName = "日历视图",
                VisibleFields = new List<string>(allNames),
                CalendarConfig = new CalendarConfig
                {
                    DateField = dateFields[0].Name,
                    TitleField = firstName.Length > 0 ? firstName : dateFields[0].Name
                }
            });
        }

        // ── 甘特视图（至少需要一个日期字段）────────────────────
        if (dateFields.Count >= 1)
        {
            config.Views.Add(new ViewConfig
            {
                ViewId = ViewConfig.NewId("gantt"),
                ViewType = ViewType.Gantt,
                ViewName = "甘特视图",
                VisibleFields = new List<string>(allNames),
                GanttConfig = new GanttConfig
                {
                    StartField = dateFields[0].Name,
                    EndField = dateFields.Count > 1 ? dateFields[1].Name : dateFields[0].Name,
                    LabelField = firstName.Length > 0 ? firstName : dateFields[0].Name
                }
            });
        }

        // ── 仪表盘 ──────────────────────────────────────────────
        config.Views.Add(new ViewConfig
        {
            ViewId = ViewConfig.NewId("dashboard"),
            ViewType = ViewType.Dashboard,
            ViewName = "仪表盘",
            VisibleFields = new List<string>(allNames),
            DashboardConfig = BuildDefaultDashboard(table)
        });

        // ── 统计图表 ────────────────────────────────────────────
        config.Views.Add(new ViewConfig
        {
            ViewId = ViewConfig.NewId("chart"),
            ViewType = ViewType.Chart,
            ViewName = "统计图表",
            VisibleFields = new List<string>(allNames),
            ChartConfig = BuildDefaultChart(table)
        });

        // ── 自动识别 数量 × 单价 = 金额 联动 ────────────────────
        var link = DetectNumericLink(table);
        if (link != null) config.NumericLinks.Add(link);

        return config;
    }

    private static CardMeta BuildCardMeta(DataTableModel table)
    {
        var imageField = table.Fields.Find(f =>
            f.Type == FieldType.Image ||
            f.Name.Contains("图片") || f.Name.Contains("照片") ||
            f.Name.Contains("image", StringComparison.OrdinalIgnoreCase));

        return new CardMeta
        {
            Title = table.Fields.Count > 0 ? table.Fields[0].Name : string.Empty,
            Image = imageField?.Name ?? string.Empty,
            Description = table.Fields.Skip(1).Take(4).Select(f => f.Name).ToList()
        };
    }

    /// <summary>依据字段类型自动搭一个可用的仪表盘</summary>
    public static DashboardConfig BuildDefaultDashboard(DataTableModel table)
    {
        var dash = new DashboardConfig { Columns = 2 };

        var numericFields = table.Fields.Where(f => FieldTypeHelper.IsNumeric(f.Type)).ToList();
        var moneyField = numericFields.Find(f => f.Type == FieldType.Currency);
        var qtyField = numericFields.Find(f => f.Type == FieldType.Integer);
        var dimFields = table.Fields.Where(f => FieldTypeHelper.IsDimension(f.Type)).ToList();
        var timeField = table.Fields.Find(f => FieldTypeHelper.IsTemporal(f.Type));

        // KPI 卡片：总记录数 + 金额合计 + 数量合计 + 主维度去重数
        dash.StatCards.Add(new StatCardConfig
        {
            Id = ViewConfig.NewId("kpi"),
            Title = "记录总数",
            Field = table.Fields.Count > 0 ? table.Fields[0].Name : string.Empty,
            Aggregation = AggregateMode.Count,
            Format = "int",
            Color = "#4E7CF6"
        });

        if (moneyField != null)
        {
            dash.StatCards.Add(new StatCardConfig
            {
                Id = ViewConfig.NewId("kpi"),
                Title = moneyField.Name + "合计",
                Field = moneyField.Name,
                Aggregation = AggregateMode.Sum,
                Format = "money",
                Color = "#F2994A"
            });
            dash.StatCards.Add(new StatCardConfig
            {
                Id = ViewConfig.NewId("kpi"),
                Title = moneyField.Name + "均值",
                Field = moneyField.Name,
                Aggregation = AggregateMode.Average,
                Format = "money",
                Color = "#27AE60"
            });
        }

        if (qtyField != null)
        {
            dash.StatCards.Add(new StatCardConfig
            {
                Id = ViewConfig.NewId("kpi"),
                Title = qtyField.Name + "合计",
                Field = qtyField.Name,
                Aggregation = AggregateMode.Sum,
                Format = "int",
                Color = "#9B51E0"
            });
        }

        if (dimFields.Count > 0 && dash.StatCards.Count < 4)
        {
            dash.StatCards.Add(new StatCardConfig
            {
                Id = ViewConfig.NewId("kpi"),
                Title = dimFields[0].Name + "数",
                Field = dimFields[0].Name,
                Aggregation = AggregateMode.DistinctCount,
                Format = "int",
                Color = "#2D9CDB"
            });
        }

        string metric = moneyField?.Name ?? qtyField?.Name ?? string.Empty;
        var metricAgg = metric.Length > 0 ? AggregateMode.Sum : AggregateMode.Count;
        string countField = table.Fields.Count > 0 ? table.Fields[0].Name : string.Empty;

        // 图 1：主维度分布（柱状图）
        if (dimFields.Count > 0)
        {
            dash.Charts.Add(new ChartConfig
            {
                Id = ViewConfig.NewId("chart"),
                Title = dimFields[0].Name + (metric.Length > 0 ? " · " + metric + "合计" : " · 记录数"),
                Type = ChartType.Column,
                DimensionField = dimFields[0].Name,
                MetricField = metric.Length > 0 ? metric : countField,
                Aggregation = metricAgg
            });
        }

        // 图 2：时间趋势（按月折线）
        if (timeField != null)
        {
            dash.Charts.Add(new ChartConfig
            {
                Id = ViewConfig.NewId("chart"),
                Title = "趋势 · 按月",
                Type = ChartType.Line,
                TimeField = timeField.Name,
                TimeGroup = TimeDimension.Month,
                MetricField = metric.Length > 0 ? metric : countField,
                Aggregation = metricAgg
            });

            // 图 3：季度对比（柱状图）
            dash.Charts.Add(new ChartConfig
            {
                Id = ViewConfig.NewId("chart"),
                Title = "季度对比",
                Type = ChartType.Column,
                TimeField = timeField.Name,
                TimeGroup = TimeDimension.Quarter,
                MetricField = metric.Length > 0 ? metric : countField,
                Aggregation = metricAgg
            });
        }

        // 图 4：次维度占比（环形图）
        var secondDim = dimFields.Count > 1 ? dimFields[1] : (dimFields.Count > 0 ? dimFields[0] : null);
        if (secondDim != null)
        {
            dash.Charts.Add(new ChartConfig
            {
                Id = ViewConfig.NewId("chart"),
                Title = secondDim.Name + "占比",
                Type = ChartType.Doughnut,
                DimensionField = secondDim.Name,
                MetricField = metric.Length > 0 ? metric : countField,
                Aggregation = metricAgg,
                TopN = 8
            });
        }

        return dash;
    }

    /// <summary>默认单图表配置</summary>
    public static ChartConfig BuildDefaultChart(DataTableModel table)
    {
        var metricField = table.Fields.Find(f => f.Type == FieldType.Currency)
                       ?? table.Fields.Find(f => FieldTypeHelper.IsNumeric(f.Type));
        var dimField = table.Fields.Find(f => FieldTypeHelper.IsDimension(f.Type));
        var timeField = table.Fields.Find(f => FieldTypeHelper.IsTemporal(f.Type));
        string countField = table.Fields.Count > 0 ? table.Fields[0].Name : string.Empty;

        var chart = new ChartConfig
        {
            Id = ViewConfig.NewId("chart"),
            Title = "统计图表",
            Type = ChartType.Column,
            MetricField = metricField?.Name ?? countField,
            Aggregation = metricField != null ? AggregateMode.Sum : AggregateMode.Count,
            Height = 420
        };

        if (timeField != null)
        {
            chart.TimeField = timeField.Name;
            chart.TimeGroup = TimeDimension.Quarter;
            chart.Title = "季度统计";
        }
        else if (dimField != null)
        {
            chart.DimensionField = dimField.Name;
            chart.Title = dimField.Name + "统计";
        }

        return chart;
    }

    /// <summary>自动探测 数量 × 单价 = 金额 三字段联动</summary>
    public static NumericLinkConfig? DetectNumericLink(DataTableModel table)
    {
        string? Find(params string[] keys) => table.Fields
            .FirstOrDefault(f => keys.Any(k => f.Name.Contains(k, StringComparison.OrdinalIgnoreCase)))?.Name;

        string? qty = Find("数量", "件数", "台数", "数目");
        string? price = Find("单价", "价格", "售价", "单位价");
        string? amount = Find("金额", "总价", "总额", "合计金额");

        if (qty == null || price == null || amount == null) return null;
        if (qty == price || price == amount || qty == amount) return null;

        return new NumericLinkConfig
        {
            QuantityField = qty,
            UnitPriceField = price,
            AmountField = amount
        };
    }
}
