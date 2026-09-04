using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MultiTableAddin.Core;
using TextBox = System.Windows.Controls.TextBox;
using UserControl = System.Windows.Controls.UserControl;
using System.Windows.Input;

namespace MultiTableAddin.Views.Controls;

/// <summary>
/// 图表配置编辑面板。
/// ChartView 内嵌使用，仪表盘通过 ChartConfigDialog 弹窗使用，避免两处重复实现。
/// 所有修改直接写回传入的 ChartConfig 实例，并抛出 Changed 事件供宿主重绘。
/// </summary>
public class ChartConfigEditor : UserControl
{
    private const string NoneText = "（不设置）";

    private readonly StackPanel _root = new() { Margin = new Thickness(0) };

    private DataTableModel? _table;
    private ViewConfigFile? _config;
    private ChartConfig? _cfg;
    private bool _loading;

    /// <summary>配置发生变化</summary>
    public event EventHandler? Changed;

    public ChartConfig? Config => _cfg;

    public ChartConfigEditor()
    {
        Content = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = _root
        };
    }

    public void Load(DataTableModel table, ViewConfigFile? config, ChartConfig cfg)
    {
        _table = table;
        _config = config;
        _cfg = cfg;
        Rebuild();
    }

    private void Raise() => Changed?.Invoke(this, EventArgs.Empty);

    // ── 字段集合 ──────────────────────────────────────────────

    private List<string> AllFields =>
        _table?.Fields.Select(f => f.Name).ToList() ?? new List<string>();

    private FieldType TypeOf(string name)
    {
        if (_config != null) return _config.GetEffectiveFieldType(name);
        return _table?.FindField(name)?.Type ?? FieldType.Text;
    }

    private List<string> DateFields =>
        AllFields.Where(n => FieldTypeHelper.IsTemporal(TypeOf(n))).ToList();

    private List<string> NumericFields =>
        AllFields.Where(n => FieldTypeHelper.IsNumeric(TypeOf(n))).ToList();

    private List<string> DimensionFields
    {
        get
        {
            // 维度优先展示离散型字段，其余字段追加在后面，避免用户找不到列
            var dims = AllFields.Where(n => FieldTypeHelper.IsDimension(TypeOf(n))).ToList();
            var rest = AllFields.Where(n => !dims.Contains(n)).ToList();
            dims.AddRange(rest);
            return dims;
        }
    }

    private List<string> MetricFields
    {
        get
        {
            var nums = NumericFields;
            var rest = AllFields.Where(n => !nums.Contains(n)).ToList();
            nums.AddRange(rest);
            return nums;
        }
    }

    // ── 构建 ─────────────────────────────────────────────────

    private void Rebuild()
    {
        _loading = true;
        try
        {
            _root.Children.Clear();
            if (_cfg == null || _table == null) return;

            var cfg = _cfg;

            AddRow("标题", TextBoxFor(cfg.Title, v => cfg.Title = v));

            AddRow("图表类型", EnumCombo(ChartTypeHelper.AllLabels, cfg.Type, v =>
            {
                cfg.Type = v;
                Rebuild();   // 类型变化后可编辑项不同，整体重建
            }));

            bool isGauge = cfg.Type == ChartType.Gauge;
            bool supportSeries = cfg.Type is ChartType.Column or ChartType.Bar or ChartType.Line or ChartType.Area;

            if (!isGauge)
            {
                AddRow("分组维度", FieldCombo(DimensionFields, cfg.DimensionField, v => cfg.DimensionField = v));
                AddRow("时间字段", FieldCombo(DateFields, cfg.TimeField, v => cfg.TimeField = v));
                AddRow("时间粒度", EnumCombo(TimeDimensionHelper.AllLabels, cfg.TimeGroup, v => cfg.TimeGroup = v));
                AddHint("设置「时间字段 + 时间粒度」后，将按年 / 季度 / 月 / 周 / 日 汇总，忽略分组维度。");
            }

            AddRow("度量字段", FieldCombo(MetricFields, cfg.MetricField, v => cfg.MetricField = v));
            AddRow("聚合方式", EnumCombo(AggregateModeHelper.AllLabels, cfg.Aggregation, v => cfg.Aggregation = v));

            if (supportSeries)
            {
                AddRow("次级系列", FieldCombo(DimensionFields, cfg.SeriesField, v => cfg.SeriesField = v));
            }

            if (!isGauge)
            {
                AddRow("最多分类", NumberBox(cfg.TopN, v => cfg.TopN = (int)Math.Max(0, v), "0"));
            }
            else
            {
                AddRow("目标值", NumberBox(cfg.GaugeTarget, v => cfg.GaugeTarget = v));
            }

            AddRow("图表高度", NumberBox(cfg.Height, v => cfg.Height = Math.Max(140, v), "0"));
            AddRow("占据列数", SpanCombo(cfg.ColumnSpan, v => cfg.ColumnSpan = v));
        }
        finally
        {
            _loading = false;
        }
    }

    private void AddRow(string label, FrameworkElement editor)
    {
        var grid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var text = new TextBlock
        {
            Text = label,
            FontSize = 12,
            TextAlignment = TextAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 0)
        };
        text.SetResourceReference(TextBlock.ForegroundProperty, "SecondaryTextBrush");

        Grid.SetColumn(text, 0);
        Grid.SetColumn(editor, 1);
        grid.Children.Add(text);
        grid.Children.Add(editor);
        _root.Children.Add(grid);
    }

    private void AddHint(string text)
    {
        var tb = new TextBlock
        {
            Text = text,
            FontSize = 11,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(72 + 8, -4, 0, 10)
        };
        tb.SetResourceReference(TextBlock.ForegroundProperty, "ThirdlyTextBrush");
        _root.Children.Add(tb);
    }

    // ── 控件工厂 ──────────────────────────────────────────────

    private TextBox TextBoxFor(string value, Action<string> setter)
    {
        var tb = new TextBox { Text = value ?? string.Empty, FontSize = 12 };
        tb.TextChanged += (_, _) =>
        {
            if (_loading) return;
            setter(tb.Text);
            Raise();
        };
        return tb;
    }

    private ComboBox FieldCombo(IEnumerable<string> names, string current, Action<string> setter)
    {
        var cb = new ComboBox { FontSize = 12 };
        cb.Items.Add(NoneText);
        foreach (var n in names) cb.Items.Add(n);

        if (!string.IsNullOrEmpty(current) && cb.Items.Contains(current))
            cb.SelectedItem = current;
        else
            cb.SelectedIndex = 0;

        cb.SelectionChanged += (_, _) =>
        {
            if (_loading) return;
            string s = cb.SelectedItem as string ?? string.Empty;
            setter(s == NoneText ? string.Empty : s);
            Raise();
        };
        return cb;
    }

    private ComboBox EnumCombo<T>(IEnumerable<KeyValuePair<T, string>> items, T current, Action<T> setter)
        where T : struct
    {
        var cb = new ComboBox
        {
            FontSize = 12,
            DisplayMemberPath = "Value",
            SelectedValuePath = "Key"
        };
        foreach (var kv in items) cb.Items.Add(kv);
        cb.SelectedValue = current;
        if (cb.SelectedIndex < 0 && cb.Items.Count > 0) cb.SelectedIndex = 0;

        cb.SelectionChanged += (_, _) =>
        {
            if (_loading) return;
            if (cb.SelectedValue is T v)
            {
                setter(v);
                Raise();
            }
        };
        return cb;
    }

    private ComboBox SpanCombo(int current, Action<int> setter)
    {
        var cb = new ComboBox { FontSize = 12 };
        cb.Items.Add("1 列");
        cb.Items.Add("2 列（整行）");
        cb.SelectedIndex = current >= 2 ? 1 : 0;
        cb.SelectionChanged += (_, _) =>
        {
            if (_loading) return;
            setter(cb.SelectedIndex >= 1 ? 2 : 1);
            Raise();
        };
        return cb;
    }

    private TextBox NumberBox(double value, Action<double> setter, string format = "0.####")
    {
        double original = value;
        var tb = new TextBox
        {
            Text = value.ToString(format, CultureInfo.InvariantCulture),
            FontSize = 12
        };
        tb.LostFocus += (_, _) =>
        {
            if (_loading) return;
            if (double.TryParse(tb.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
            {
                setter(v);
                Raise();
            }
            else
            {
                tb.Text = original.ToString(format, CultureInfo.InvariantCulture);
            }
        };
        tb.KeyDown += (_, e) =>
        {
            if (e.Key == System.Windows.Input.Key.Enter)
                (System.Windows.Input.Keyboard.FocusedElement as FrameworkElement)?
                    .MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        };
        return tb;
    }
}
