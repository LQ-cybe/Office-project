using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MultiTableAddin.Core;
using MultiTableAddin.Views.Controls;
using UserControl = System.Windows.Controls.UserControl;

namespace MultiTableAddin.Views;

/// <summary>
/// 单图表视图：左侧大图 + 右侧配置面板，顶部提供时间粒度 / 日期区间 / 关键字快捷筛选。
/// 配置直接写回 ViewConfig.ChartConfig，随视图配置一起保存。
/// </summary>
public partial class ChartView : UserControl, ITableView, IConfigAware
{
    private static readonly (string Label, TimeDimension? Dim)[] QuickTimeOptions =
    {
        ("按配置", null),
        ("按年", TimeDimension.Year),
        ("按季度", TimeDimension.Quarter),
        ("按月", TimeDimension.Month),
        ("按周", TimeDimension.Week),
        ("按日", TimeDimension.Day),
        ("不按时间", TimeDimension.None)
    };

    private readonly ChartConfigEditor _editor = new();

    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;
    private ViewConfigFile? _configFile;

    private List<DataRowModel> _allRows = new();
    private bool _loaded;
    private bool _suspend;
    private bool _panelVisible = true;

    public ChartView()
    {
        InitializeComponent();
        EditorHost.Content = _editor;
        _editor.Changed += (_, _) => Redraw();
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;
        _allRows = viewData.Groups.SelectMany(g => g.Rows).ToList();

        if (_loaded) Setup();
    }

    public void SetConfigFile(ViewConfigFile configFile)
    {
        _configFile = configFile;
        if (_loaded) Setup();
    }

    private void ChartView_Loaded(object sender, RoutedEventArgs e)
    {
        if (_loaded) return;
        _loaded = true;
        Setup();
    }

    private void Setup()
    {
        if (_dataTable == null || _viewConfig == null) return;

        _suspend = true;
        try
        {
            if (QuickTimeCombo.Items.Count == 0)
            {
                foreach (var o in QuickTimeOptions) QuickTimeCombo.Items.Add(o.Label);
                QuickTimeCombo.SelectedIndex = 0;
            }
        }
        finally
        {
            _suspend = false;
        }

        _editor.Load(_dataTable, _configFile, EnsureChartConfig());
        Redraw();
    }

    private ChartConfig EnsureChartConfig()
    {
        var cfg = _viewConfig!.ChartConfig;
        if (cfg == null)
        {
            cfg = new ChartConfig
            {
                Id = ViewConfig.NewId("chart"),
                Title = _viewConfig.ViewName,
                Type = ChartType.Column,
                Aggregation = AggregateMode.Count,
                Height = 320
            };

            if (_dataTable != null)
            {
                cfg.DimensionField = _dataTable.Fields
                    .FirstOrDefault(f => FieldTypeHelper.IsDimension(EffectiveType(f.Name)))?.Name ?? string.Empty;
                var metric = _dataTable.Fields
                    .FirstOrDefault(f => FieldTypeHelper.IsNumeric(EffectiveType(f.Name)))?.Name;
                if (!string.IsNullOrEmpty(metric))
                {
                    cfg.MetricField = metric!;
                    cfg.Aggregation = AggregateMode.Sum;
                }
            }

            _viewConfig.ChartConfig = cfg;
        }
        return cfg;
    }

    private FieldType EffectiveType(string name) =>
        _configFile?.GetEffectiveFieldType(name) ?? _dataTable?.FindField(name)?.Type ?? FieldType.Text;

    // ── 筛选 ─────────────────────────────────────────────────

    private List<DataRowModel> FilteredRows()
    {
        IEnumerable<DataRowModel> rows = _allRows;

        var cfg = _viewConfig?.ChartConfig;
        string timeField = cfg?.TimeField ?? string.Empty;
        if (string.IsNullOrEmpty(timeField))
        {
            timeField = _dataTable?.Fields
                .FirstOrDefault(f => FieldTypeHelper.IsTemporal(EffectiveType(f.Name)))?.Name ?? string.Empty;
        }

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

    private ChartConfig EffectiveConfig()
    {
        var cfg = EnsureChartConfig();
        int i = QuickTimeCombo.SelectedIndex;
        TimeDimension? dim = i >= 0 && i < QuickTimeOptions.Length ? QuickTimeOptions[i].Dim : null;
        if (dim == null || cfg.Type == ChartType.Gauge) return cfg;

        var c = cfg.Clone();
        c.TimeGroup = dim.Value;
        if (dim.Value != TimeDimension.None && string.IsNullOrEmpty(c.TimeField))
        {
            c.TimeField = _dataTable?.Fields
                .FirstOrDefault(f => FieldTypeHelper.IsTemporal(EffectiveType(f.Name)))?.Name ?? string.Empty;
            if (string.IsNullOrEmpty(c.TimeField)) return cfg;
        }
        return c;
    }

    // ── 绘制 ─────────────────────────────────────────────────

    private void Redraw()
    {
        if (_dataTable == null || _viewConfig == null || !_loaded) return;

        var rows = FilteredRows();
        var effective = EffectiveConfig();

        RowSummary.Text = rows.Count == _allRows.Count
            ? $"共 {_allRows.Count} 行"
            : $"筛选 {rows.Count} / {_allRows.Count} 行";

        ChartDataSet ds;
        try
        {
            ds = ChartDataBuilder.Build(rows, effective);
        }
        catch (Exception ex)
        {
            ds = new ChartDataSet { Message = "聚合失败: " + ex.Message };
            AddInLog.Write("ChartView.Build.Error", ex.ToString());
        }

        ChartTitle.Text = string.IsNullOrWhiteSpace(effective.Title) ? _viewConfig.ViewName : effective.Title;
        ChartSubtitle.Text = Describe(effective);
        ChartMessage.Text = ds.Message;

        Chart.ChartType = effective.Type;
        Chart.GaugeTarget = effective.GaugeTarget;
        Chart.Data = ds;
    }

    private static string Describe(ChartConfig cfg)
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

    // ── 事件 ─────────────────────────────────────────────────

    private void OnQuickChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suspend || !_loaded) return;
        Redraw();
    }

    private void OnSearchChanged(object sender, TextChangedEventArgs e)
    {
        if (_suspend || !_loaded) return;
        Redraw();
    }

    private void OnResetClick(object sender, RoutedEventArgs e)
    {
        _suspend = true;
        try
        {
            StartPicker.SelectedDate = null;
            EndPicker.SelectedDate = null;
            SearchBox.Text = string.Empty;
            QuickTimeCombo.SelectedIndex = 0;
        }
        finally
        {
            _suspend = false;
        }
        Redraw();
    }

    private void OnTogglePanelClick(object sender, RoutedEventArgs e)
    {
        _panelVisible = !_panelVisible;
        PanelBorder.Visibility = _panelVisible ? Visibility.Visible : Visibility.Collapsed;
        PanelColumn.Width = _panelVisible ? new GridLength(272) : new GridLength(0);
        BtnTogglePanel.Content = _panelVisible ? "隐藏配置" : "显示配置";
    }
}
