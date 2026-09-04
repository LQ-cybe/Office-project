using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MultiTableAddin.Core;
using MultiTableAddin.Views.Charts;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;
using ColorConverter = System.Windows.Media.ColorConverter;
using Pen = System.Windows.Media.Pen;
using Size = System.Windows.Size;
using FontFamily = System.Windows.Media.FontFamily;

namespace MultiTableAddin.Views;

/// <summary>
/// 仪表盘汇总视图：顶部 KPI 指标卡 + 下方多图表网格。
/// 支持全局时间维度切换（年 / 季度 / 月 / 周 / 日）、日期区间筛选与全字段关键字搜索。
/// </summary>
public partial class DashboardView : UserControl, ITableView, IConfigAware
{
    private static readonly (string Label, TimeDimension? Dim)[] TimeGroupOptions =
    {
        ("各图表自定义", null),
        ("按年", TimeDimension.Year),
        ("按季度", TimeDimension.Quarter),
        ("按月", TimeDimension.Month),
        ("按周", TimeDimension.Week),
        ("按日", TimeDimension.Day),
        ("不按时间", TimeDimension.None)
    };

    private const string NoTimeFieldText = "（不限）";

    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;
    private ViewConfigFile? _configFile;

    private List<DataRowModel> _allRows = new();
    private bool _loaded;
    private bool _suspend;

    public DashboardView()
    {
        InitializeComponent();
    }

    // ── 接口实现 ─────────────────────────────────────────────

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;
        _allRows = viewData.Groups.SelectMany(g => g.Rows).ToList();

        if (_loaded) InitToolbarAndBuild();
    }

    public void SetConfigFile(ViewConfigFile configFile)
    {
        _configFile = configFile;
        if (_loaded) InitToolbarAndBuild();
    }

    private void DashboardView_Loaded(object sender, RoutedEventArgs e)
    {
        if (_loaded) return;
        _loaded = true;
        InitToolbarAndBuild();
    }

    // ── 工具条 ───────────────────────────────────────────────

    private void InitToolbarAndBuild()
    {
        if (_dataTable == null || _viewConfig == null) return;

        _suspend = true;
        try
        {
            string keepField = TimeFieldCombo.SelectedItem as string ?? string.Empty;
            int keepGroup = TimeGroupCombo.SelectedIndex;

            TimeFieldCombo.Items.Clear();
            TimeFieldCombo.Items.Add(NoTimeFieldText);
            foreach (var f in DateFields()) TimeFieldCombo.Items.Add(f);

            if (!string.IsNullOrEmpty(keepField) && TimeFieldCombo.Items.Contains(keepField))
                TimeFieldCombo.SelectedItem = keepField;
            else
                TimeFieldCombo.SelectedIndex = TimeFieldCombo.Items.Count > 1 ? 1 : 0;

            if (TimeGroupCombo.Items.Count == 0)
            {
                foreach (var opt in TimeGroupOptions) TimeGroupCombo.Items.Add(opt.Label);
                TimeGroupCombo.SelectedIndex = 0;
            }
            else if (keepGroup >= 0)
            {
                TimeGroupCombo.SelectedIndex = keepGroup;
            }
        }
        finally
        {
            _suspend = false;
        }

        BuildAll();
    }

    private List<string> DateFields()
    {
        if (_dataTable == null) return new List<string>();
        return _dataTable.Fields
            .Where(f => FieldTypeHelper.IsTemporal(EffectiveType(f.Name)))
            .Select(f => f.Name)
            .ToList();
    }

    private FieldType EffectiveType(string name) =>
        _configFile?.GetEffectiveFieldType(name) ?? _dataTable?.FindField(name)?.Type ?? FieldType.Text;

    private string SelectedTimeField
    {
        get
        {
            string s = TimeFieldCombo.SelectedItem as string ?? string.Empty;
            return s == NoTimeFieldText ? string.Empty : s;
        }
    }

    private TimeDimension? SelectedTimeGroup
    {
        get
        {
            int i = TimeGroupCombo.SelectedIndex;
            return i >= 0 && i < TimeGroupOptions.Length ? TimeGroupOptions[i].Dim : null;
        }
    }

    private void OnFilterChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suspend || !_loaded) return;
        BuildAll();
    }

    private void OnSearchChanged(object sender, TextChangedEventArgs e)
    {
        if (_suspend || !_loaded) return;
        BuildAll();
    }

    private void OnResetFilterClick(object sender, RoutedEventArgs e)
    {
        _suspend = true;
        try
        {
            StartPicker.SelectedDate = null;
            EndPicker.SelectedDate = null;
            SearchBox.Text = string.Empty;
            TimeGroupCombo.SelectedIndex = 0;
        }
        finally
        {
            _suspend = false;
        }
        BuildAll();
    }

    // ── 数据筛选 ─────────────────────────────────────────────

    private List<DataRowModel> FilteredRows()
    {
        IEnumerable<DataRowModel> rows = _allRows;

        string timeField = SelectedTimeField;
        DateTime? start = StartPicker.SelectedDate;
        DateTime? end = EndPicker.SelectedDate;

        if (!string.IsNullOrEmpty(timeField) && (start.HasValue || end.HasValue))
        {
            DateTime lo = start ?? DateTime.MinValue;
            DateTime hi = (end ?? DateTime.MaxValue.Date).Date.AddDays(1).AddTicks(-1);
            rows = rows.Where(r =>
                ValueFormatter.TryToDateTime(r.GetValue(timeField), out DateTime dt) && dt >= lo && dt <= hi);
        }

        string keyword = SearchBox.Text?.Trim() ?? string.Empty;
        if (keyword.Length > 0)
        {
            rows = rows.Where(r => r.Values.Values.Any(v =>
                ValueFormatter.ToDisplayText(v).IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0));
        }

        return rows.ToList();
    }

    // ── 主体构建 ─────────────────────────────────────────────

    private void BuildAll()
    {
        if (_dataTable == null || _viewConfig == null) return;

        var dash = EnsureDashboardConfig();
        var rows = FilteredRows();

        RowSummary.Text = rows.Count == _allRows.Count
            ? $"共 {_allRows.Count} 行"
            : $"筛选 {rows.Count} / {_allRows.Count} 行";

        BuildStats(dash, rows);
        BuildCharts(dash, rows);

        EmptyHint.Text = dash.StatCards.Count == 0 && dash.Charts.Count == 0
            ? "当前仪表盘还没有内容。点击右上角「添加指标」或「添加图表」开始搭建。"
            : string.Empty;
    }

    private DashboardConfig EnsureDashboardConfig()
    {
        _viewConfig!.DashboardConfig ??= new DashboardConfig();
        if (_viewConfig.DashboardConfig.Columns <= 0) _viewConfig.DashboardConfig.Columns = 2;
        return _viewConfig.DashboardConfig;
    }

    private void BuildStats(DashboardConfig dash, IReadOnlyList<DataRowModel> rows)
    {
        StatPanel.Children.Clear();
        for (int i = 0; i < dash.StatCards.Count; i++)
        {
            StatPanel.Children.Add(CreateStatCard(dash, dash.StatCards[i], rows, i));
        }
    }

    private Border CreateStatCard(DashboardConfig dash, StatCardConfig cfg, IReadOnlyList<DataRowModel> rows, int index)
    {
        double value;
        try { value = ChartDataBuilder.BuildStat(rows, cfg); }
        catch { value = 0; }

        string accent = string.IsNullOrWhiteSpace(cfg.Color) ? ChartDataBuilder.ColorAt(index) : cfg.Color!;

        var card = new Border
        {
            Width = 196,
            Height = 88,
            CornerRadius = new CornerRadius(8),
            Margin = new Thickness(0, 0, 12, 12),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(0)
        };
        card.SetResourceReference(Border.BackgroundProperty, "SecondaryRegionBrush");
        card.SetResourceReference(Border.BorderBrushProperty, "BorderBrush");

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(4) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var bar = new Border
        {
            Background = SafeBrush(accent),
            CornerRadius = new CornerRadius(8, 0, 0, 8)
        };
        Grid.SetColumn(bar, 0);
        grid.Children.Add(bar);

        var stack = new StackPanel { Margin = new Thickness(12, 12, 12, 10) };

        var title = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(cfg.Title) ? cfg.Field : cfg.Title,
            FontSize = 12,
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        title.SetResourceReference(TextBlock.ForegroundProperty, "SecondaryTextBrush");

        var valueText = new TextBlock
        {
            Text = ValueFormatter.ToCompactNumber(value, cfg.Format ?? "auto"),
            FontSize = 24,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = SafeBrush(accent),
            TextTrimming = TextTrimming.CharacterEllipsis
        };

        var sub = new TextBlock
        {
            Text = AggregateModeHelper.GetLabel(cfg.Aggregation) +
                   (string.IsNullOrEmpty(cfg.Field) ? string.Empty : " · " + cfg.Field),
            FontSize = 11,
            Margin = new Thickness(0, 2, 0, 0),
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        sub.SetResourceReference(TextBlock.ForegroundProperty, "ThirdlyTextBrush");

        stack.Children.Add(title);
        stack.Children.Add(valueText);
        stack.Children.Add(sub);
        Grid.SetColumn(stack, 1);
        grid.Children.Add(stack);

        card.Child = grid;
        card.ToolTip = $"{cfg.Title}\n{AggregateModeHelper.GetLabel(cfg.Aggregation)}({cfg.Field})\n原始值: {value:0.####}"
                       + (string.IsNullOrWhiteSpace(cfg.Filter) ? string.Empty : "\n筛选: " + cfg.Filter);

        var menu = new ContextMenu();
        var edit = new MenuItem { Header = "编辑指标" };
        edit.Click += (_, _) => EditStatCard(dash, cfg);
        var del = new MenuItem { Header = "删除指标" };
        del.Click += (_, _) =>
        {
            dash.StatCards.Remove(cfg);
            BuildAll();
        };
        menu.Items.Add(edit);
        menu.Items.Add(del);
        card.ContextMenu = menu;
        card.MouseLeftButtonUp += (_, e) =>
        {
            if (e.ClickCount == 2) EditStatCard(dash, cfg);
        };

        return card;
    }

    private void EditStatCard(DashboardConfig dash, StatCardConfig cfg)
    {
        if (_dataTable == null) return;
        var dlg = new StatCardDialog(_dataTable, _configFile, cfg) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            int idx = dash.StatCards.IndexOf(cfg);
            if (idx >= 0) dash.StatCards[idx] = dlg.Result;
            BuildAll();
        }
    }

    private void BuildCharts(DashboardConfig dash, IReadOnlyList<DataRowModel> rows)
    {
        ChartGrid.Children.Clear();
        ChartGrid.ColumnDefinitions.Clear();
        ChartGrid.RowDefinitions.Clear();

        int cols = Math.Max(1, dash.Columns);
        for (int i = 0; i < cols; i++)
            ChartGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        int r = 0, c = 0;
        foreach (var cfg in dash.Charts.ToList())
        {
            int span = Math.Min(Math.Max(1, cfg.ColumnSpan), cols);
            if (c + span > cols) { r++; c = 0; }
            while (ChartGrid.RowDefinitions.Count <= r)
                ChartGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var card = CreateChartCard(dash, cfg, rows);
            Grid.SetRow(card, r);
            Grid.SetColumn(card, c);
            Grid.SetColumnSpan(card, span);
            ChartGrid.Children.Add(card);

            c += span;
            if (c >= cols) { c = 0; r++; }
        }
    }

    /// <summary>把全局时间维度选择叠加到单个图表配置上</summary>
    private ChartConfig EffectiveChart(ChartConfig cfg)
    {
        var dim = SelectedTimeGroup;
        if (dim == null || cfg.Type == ChartType.Gauge) return cfg;

        var c = cfg.Clone();
        c.TimeGroup = dim.Value;
        if (dim.Value != TimeDimension.None)
        {
            string tf = SelectedTimeField;
            if (!string.IsNullOrEmpty(tf)) c.TimeField = tf;
            if (string.IsNullOrEmpty(c.TimeField)) return cfg;   // 没有可用时间字段则保持原配置
        }
        return c;
    }

    private Border CreateChartCard(DashboardConfig dash, ChartConfig cfg, IReadOnlyList<DataRowModel> rows)
    {
        var effective = EffectiveChart(cfg);

        ChartDataSet ds;
        try { ds = ChartDataBuilder.Build(rows, effective); }
        catch (Exception ex) { ds = new ChartDataSet { Title = cfg.Title, Message = "聚合失败: " + ex.Message }; }

        var card = new Border
        {
            CornerRadius = new CornerRadius(8),
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0, 0, 12, 12),
            Padding = new Thickness(12, 10, 12, 12)
        };
        card.SetResourceReference(Border.BackgroundProperty, "SecondaryRegionBrush");
        card.SetResourceReference(Border.BorderBrushProperty, "BorderBrush");

        var dock = new DockPanel();

        // 标题栏
        var header = new DockPanel { Margin = new Thickness(0, 0, 0, 8) };
        DockPanel.SetDock(header, Dock.Top);

        var buttons = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal };
        DockPanel.SetDock(buttons, Dock.Right);
        buttons.Children.Add(MiniButton("设置", () => EditChart(dash, cfg)));
        buttons.Children.Add(MiniButton("删除", () =>
        {
            if (MessageBox.Show($"确定删除图表「{cfg.Title}」？", "删除图表",
                    MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK) return;
            dash.Charts.Remove(cfg);
            BuildAll();
        }));

        var title = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(cfg.Title) ? "未命名图表" : cfg.Title,
            FontSize = 13,
            FontWeight = FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        title.SetResourceReference(TextBlock.ForegroundProperty, "PrimaryTextBrush");

        header.Children.Add(buttons);
        header.Children.Add(title);
        dock.Children.Add(header);

        // 副标题（当前生效的聚合口径）
        var subtitle = new TextBlock
        {
            Text = DescribeChart(effective),
            FontSize = 11,
            Margin = new Thickness(0, 0, 0, 6),
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        subtitle.SetResourceReference(TextBlock.ForegroundProperty, "ThirdlyTextBrush");
        DockPanel.SetDock(subtitle, Dock.Top);
        dock.Children.Add(subtitle);

        if (!string.IsNullOrEmpty(ds.Message))
        {
            var msg = new TextBlock
            {
                Text = ds.Message,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 6, 0, 0)
            };
            msg.SetResourceReference(TextBlock.ForegroundProperty, "WarningBrush");
            DockPanel.SetDock(msg, Dock.Bottom);
            dock.Children.Add(msg);
        }

        var chart = new SimpleChart
        {
            Height = Math.Max(140, cfg.Height),
            ChartType = effective.Type,
            GaugeTarget = effective.GaugeTarget,
            Data = ds
        };
        dock.Children.Add(chart);

        card.Child = dock;
        return card;
    }

    private static string DescribeChart(ChartConfig cfg)
    {
        string metric = string.IsNullOrEmpty(cfg.MetricField)
            ? AggregateModeHelper.GetLabel(cfg.Aggregation)
            : AggregateModeHelper.GetLabel(cfg.Aggregation) + "(" + cfg.MetricField + ")";

        string dim = cfg.TimeGroup != TimeDimension.None && !string.IsNullOrEmpty(cfg.TimeField)
            ? cfg.TimeField + " · " + TimeDimensionHelper.GetLabel(cfg.TimeGroup)
            : (string.IsNullOrEmpty(cfg.DimensionField) ? "全部" : cfg.DimensionField);

        string series = string.IsNullOrEmpty(cfg.SeriesField) ? string.Empty : " · 系列 " + cfg.SeriesField;
        return ChartTypeHelper.GetLabel(cfg.Type) + " · " + dim + " · " + metric + series;
    }

    private Button MiniButton(string text, Action onClick)
    {
        var btn = new Button
        {
            Content = text,
            FontSize = 11,
            Padding = new Thickness(8, 2, 8, 2),
            MinHeight = 22,
            Height = double.NaN,
            Margin = new Thickness(6, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Cursor = System.Windows.Input.Cursors.Hand
        };
        if (TryFindResource("ButtonDefault") is Style s) btn.Style = s;
        btn.Height = double.NaN;
        btn.Click += (_, _) => onClick();
        return btn;
    }

    private void EditChart(DashboardConfig dash, ChartConfig cfg)
    {
        if (_dataTable == null) return;
        var dlg = new ChartConfigDialog(_dataTable, _configFile, cfg, FilteredRows())
        {
            Owner = Window.GetWindow(this)
        };
        if (dlg.ShowDialog() == true)
        {
            int idx = dash.Charts.IndexOf(cfg);
            if (idx >= 0) dash.Charts[idx] = dlg.Result;
            BuildAll();
        }
    }

    private void OnAddChartClick(object sender, RoutedEventArgs e)
    {
        if (_dataTable == null || _viewConfig == null)
        {
            HandyControl.Controls.Growl.WarningGlobal("请先加载数据。");
            return;
        }

        var dash = EnsureDashboardConfig();
        var draft = new ChartConfig
        {
            Id = ViewConfig.NewId("chart"),
            Title = "新图表",
            Type = ChartType.Column,
            DimensionField = _dataTable.Fields
                .FirstOrDefault(f => FieldTypeHelper.IsDimension(EffectiveType(f.Name)))?.Name ?? string.Empty,
            MetricField = _dataTable.Fields
                .FirstOrDefault(f => FieldTypeHelper.IsNumeric(EffectiveType(f.Name)))?.Name ?? string.Empty,
            Aggregation = AggregateMode.Sum,
            Height = 260
        };
        if (string.IsNullOrEmpty(draft.MetricField)) draft.Aggregation = AggregateMode.Count;

        var dlg = new ChartConfigDialog(_dataTable, _configFile, draft, FilteredRows())
        {
            Owner = Window.GetWindow(this)
        };
        if (dlg.ShowDialog() == true)
        {
            dash.Charts.Add(dlg.Result);
            BuildAll();
        }
    }

    private void OnAddStatClick(object sender, RoutedEventArgs e)
    {
        if (_dataTable == null || _viewConfig == null)
        {
            HandyControl.Controls.Growl.WarningGlobal("请先加载数据。");
            return;
        }

        var dash = EnsureDashboardConfig();
        var draft = new StatCardConfig
        {
            Id = ViewConfig.NewId("stat"),
            Title = "新指标",
            Aggregation = AggregateMode.Count,
            Format = "int"
        };

        var dlg = new StatCardDialog(_dataTable, _configFile, draft) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            dash.StatCards.Add(dlg.Result);
            BuildAll();
        }
    }

    private static Brush SafeBrush(string hex)
    {
        try
        {
            var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            b.Freeze();
            return b;
        }
        catch
        {
            return Brushes.SteelBlue;
        }
    }
}
