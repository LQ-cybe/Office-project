using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MultiTableAddin.Core;
using UserControl = System.Windows.Controls.UserControl;
using TextBox = System.Windows.Controls.TextBox;
using CheckBox = System.Windows.Controls.CheckBox;
using FontFamily = System.Windows.Media.FontFamily;
using TextBoxBase = System.Windows.Controls.Primitives.TextBoxBase;
using WpfControl = System.Windows.Controls.Control;

namespace MultiTableAddin.Views.Controls;

/// <summary>
/// 通用记录编辑器：横向「字段名 + 录入控件」布局，按字段类型给出人性化输入。
/// 表单视图与画册详情共用，保证两处交互一致。
/// 内置 数量 × 单价 = 金额 三向联动。
/// </summary>
public class RecordEditor : UserControl
{
    private sealed class FieldEditor
    {
        public string Name = string.Empty;
        public FieldType Type;
        public FrameworkElement Element = null!;
        public Func<object?> Read = () => null;
        public Action<object?> Write = _ => { };
        public object? Original;
    }

    private readonly Grid _grid = new();
    private readonly List<FieldEditor> _editors = new();
    private DataTableModel? _table;
    private ViewConfigFile? _config;
    private DataRowModel? _row;
    private bool _suppressLink;

    /// <summary>只读模式下所有控件禁用</summary>
    public bool ReadOnlyMode { get; set; }

    /// <summary>为 true 时不显示货币/百分比的前后缀符号（如 ¥、%）</summary>
    public bool HideAffixes { get; set; }

    /// <summary>统一控件宽度（>0 时所有录入控件使用相同宽度）；0 表示按原样自适应</summary>
    public double UniformControlWidth { get; set; }

    /// <summary>字段名列宽下限 / 上限</summary>
    public double MinLabelWidth { get; set; } = 78;
    public double MaxLabelWidth { get; set; } = 190;

    public event EventHandler? ValueChanged;

    public RecordEditor()
    {
        Content = _grid;
        _grid.Margin = new Thickness(0);
    }

    public DataRowModel? CurrentRow => _row;

    /// <summary>加载一条记录</summary>
    public void Load(DataTableModel table, ViewConfigFile? config, DataRowModel row, IList<string>? visibleFields)
    {
        _table = table;
        _config = config;
        _row = row;

        _grid.Children.Clear();
        _grid.RowDefinitions.Clear();
        _grid.ColumnDefinitions.Clear();
        _editors.Clear();

        var fields = (visibleFields != null && visibleFields.Count > 0)
            ? visibleFields.ToList()
            : table.FieldNames;

        fields = fields.Where(f => table.FindField(f) != null).ToList();
        if (fields.Count == 0)
        {
            _grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            _grid.ColumnDefinitions.Add(new ColumnDefinition());
            var tip = new TextBlock
            {
                Text = "当前视图没有可显示的字段",
                Foreground = Res("SecondaryTextBrush"),
                Margin = new Thickness(4)
            };
            _grid.Children.Add(tip);
            return;
        }

        double labelWidth = MeasureLabelWidth(fields);

        _grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(labelWidth) });
        _grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        int r = 0;
        foreach (var name in fields)
        {
            var schema = GetEffectiveField(name);
            _grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var label = new TextBlock
            {
                Text = name,
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 8, 10, 8),
                Foreground = Res("PrimaryTextBrush"),
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = $"{name}（{FieldTypeHelper.GetLabel(schema.Type)}）"
            };
            Grid.SetRow(label, r);
            Grid.SetColumn(label, 0);
            _grid.Children.Add(label);

            var editor = BuildEditor(schema, row.GetValue(name));
            editor.Element.Margin = new Thickness(0, 5, 2, 5);
            if (ReadOnlyMode && editor.Element is WpfControl c) c.IsEnabled = false;

            Grid.SetRow(editor.Element, r);
            Grid.SetColumn(editor.Element, 1);
            if (UniformControlWidth > 0)
            {
                editor.Element.Width = UniformControlWidth;
                editor.Element.MinWidth = UniformControlWidth;
                editor.Element.MaxWidth = UniformControlWidth;
                editor.Element.HorizontalAlignment = HorizontalAlignment.Left;
            }
            _grid.Children.Add(editor.Element);

            _editors.Add(editor);
            r++;
        }
    }

    private double MeasureLabelWidth(IEnumerable<string> fields)
    {
        double dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        if (dpi <= 0) dpi = 1;

        var tf = new Typeface(new FontFamily("Microsoft YaHei UI, Microsoft YaHei, Segoe UI"),
            FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);

        double max = 0;
        foreach (var f in fields)
        {
            var ft = new FormattedText(f, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                tf, 13, Brushes.Black, dpi);
            if (ft.Width > max) max = ft.Width;
        }
        return Math.Min(MaxLabelWidth, Math.Max(MinLabelWidth, max + 14));
    }

    private FieldSchema GetEffectiveField(string name)
    {
        if (_config != null) return _config.GetEffectiveField(name);
        return _table?.FindField(name) ?? new FieldSchema { Name = name };
    }

    private Brush? Res(string key) => TryFindResource(key) as Brush;

    // ─────────────────────────────────────────────────────────────
    // 各类型录入控件
    // ─────────────────────────────────────────────────────────────

    private FieldEditor BuildEditor(FieldSchema schema, object? value)
    {
        var fe = new FieldEditor { Name = schema.Name, Type = schema.Type, Original = value };

        switch (schema.Type)
        {
            case FieldType.Date:
            case FieldType.DateTime:
                BuildDate(fe, schema, value);
                break;

            case FieldType.Quarter:
                BuildCombo(fe, FieldSchema.QuarterOptions, value, editable: false, placeholder: "请选择季度");
                break;

            case FieldType.Select:
                BuildCombo(fe, schema.Options, value, editable: true, placeholder: "请选择或输入");
                break;

            case FieldType.Checkbox:
                BuildCheckbox(fe, value);
                break;

            case FieldType.LongText:
                BuildLongText(fe, value);
                break;

            case FieldType.Currency:
                BuildNumeric(fe, value, prefix: "¥", suffix: null, placeholder: "0.00");
                break;

            case FieldType.Percentage:
                BuildPercentage(fe, value);
                break;

            case FieldType.Integer:
                BuildNumeric(fe, value, prefix: null, suffix: null, placeholder: "0", integerOnly: true);
                break;

            case FieldType.Number:
                BuildNumeric(fe, value, prefix: null, suffix: null, placeholder: "请输入数字");
                break;

            case FieldType.Email:
                BuildText(fe, value, "name@example.com");
                break;

            case FieldType.Phone:
                BuildText(fe, value, "请输入电话号码", maxLength: 20);
                break;

            case FieldType.Url:
                BuildText(fe, value, "https://");
                break;

            case FieldType.Image:
                BuildImagePath(fe, value);
                break;

            default:
                BuildText(fe, value, "请输入");
                break;
        }

        return fe;
    }

    private static void Hook(TextBox tb, Action onChanged) => tb.TextChanged += (_, _) => onChanged();

    private void BuildText(FieldEditor fe, object? value, string placeholder, int maxLength = 0)
    {
        var tb = new TextBox
        {
            Text = ValueFormatter.ToDisplayText(value),
            Padding = new Thickness(8, 6, 8, 6),
            VerticalContentAlignment = VerticalAlignment.Center
        };
        if (maxLength > 0) tb.MaxLength = maxLength;
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, placeholder);

        fe.Element = tb;
        fe.Read = () => string.IsNullOrWhiteSpace(tb.Text) ? null : tb.Text.Trim();
        fe.Write = v => tb.Text = ValueFormatter.ToDisplayText(v);
        Hook(tb, RaiseChanged);
    }

    private void BuildLongText(FieldEditor fe, object? value)
    {
        var tb = new TextBox
        {
            Text = ValueFormatter.ToDisplayText(value),
            Padding = new Thickness(8, 6, 8, 6),
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 68,
            MaxHeight = 160,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "请输入内容");

        fe.Element = tb;
        fe.Read = () => string.IsNullOrWhiteSpace(tb.Text) ? null : tb.Text;
        fe.Write = v => tb.Text = ValueFormatter.ToDisplayText(v);
        Hook(tb, RaiseChanged);
    }

    private void BuildNumeric(FieldEditor fe, object? value, string? prefix, string? suffix,
        string placeholder, bool integerOnly = false)
    {
        var tb = new TextBox
        {
            Padding = new Thickness(8, 6, 8, 6),
            VerticalContentAlignment = VerticalAlignment.Center,
            Text = ValueFormatter.TryToDouble(value, out double d)
                ? (integerOnly
                    ? Math.Round(d).ToString("0", CultureInfo.InvariantCulture)
                    : d.ToString("0.####", CultureInfo.InvariantCulture))
                : ValueFormatter.ToDisplayText(value)
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, placeholder);
        tb.PreviewTextInput += (_, e) => e.Handled = !AllowNumeric(tb, e.Text, integerOnly);

        FrameworkElement host = tb;
        if (!HideAffixes && (prefix != null || suffix != null))
        {
            var dock = new DockPanel { LastChildFill = true };
            if (prefix != null)
            {
                var pre = new TextBlock
                {
                    Text = prefix,
                    FontSize = 13,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 6, 0),
                    Foreground = Res("PrimaryBrush")
                };
                DockPanel.SetDock(pre, Dock.Left);
                dock.Children.Add(pre);
            }
            if (suffix != null)
            {
                var suf = new TextBlock
                {
                    Text = suffix,
                    FontSize = 13,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 0, 0),
                    Foreground = Res("SecondaryTextBrush")
                };
                DockPanel.SetDock(suf, Dock.Right);
                dock.Children.Add(suf);
            }
            dock.Children.Add(tb);
            host = dock;
        }

        fe.Element = host;
        fe.Read = () =>
        {
            if (string.IsNullOrWhiteSpace(tb.Text)) return null;
            if (!ValueFormatter.TryToDouble(tb.Text, out double v)) return tb.Text.Trim();
            return integerOnly ? (object)(long)Math.Round(v) : v;
        };
        fe.Write = v =>
        {
            if (v == null) { tb.Text = string.Empty; return; }
            tb.Text = ValueFormatter.TryToDouble(v, out double nv)
                ? (integerOnly
                    ? Math.Round(nv).ToString("0", CultureInfo.InvariantCulture)
                    : nv.ToString("0.####", CultureInfo.InvariantCulture))
                : ValueFormatter.ToDisplayText(v);
        };
        Hook(tb, () => { OnNumericChanged(fe); RaiseChanged(); });
    }

    private void BuildPercentage(FieldEditor fe, object? value)
    {
        double shown = 0;
        if (ValueFormatter.TryToDouble(value, out double raw))
            shown = Math.Abs(raw) <= 1.000001 ? raw * 100 : raw;

        var tb = new TextBox
        {
            Padding = new Thickness(8, 6, 8, 6),
            VerticalContentAlignment = VerticalAlignment.Center,
            Text = value == null ? string.Empty : shown.ToString("0.##", CultureInfo.InvariantCulture)
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "输入 85 表示 85%");
        tb.PreviewTextInput += (_, e) => e.Handled = !AllowNumeric(tb, e.Text, false);

        var dock = new DockPanel { LastChildFill = true };
        var suf = new TextBlock
        {
            Text = "%",
            FontSize = 13,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 0, 0),
            Foreground = Res("SecondaryTextBrush")
        };
        DockPanel.SetDock(suf, Dock.Right);
        dock.Children.Add(suf);
        dock.Children.Add(tb);

        fe.Element = dock;
        fe.Read = () =>
        {
            if (string.IsNullOrWhiteSpace(tb.Text)) return null;
            return ValueFormatter.TryToDouble(tb.Text, out double v) ? v / 100.0 : (object)tb.Text.Trim();
        };
        fe.Write = v =>
        {
            if (v == null) { tb.Text = string.Empty; return; }
            if (!ValueFormatter.TryToDouble(v, out double nv)) { tb.Text = ValueFormatter.ToDisplayText(v); return; }
            double pv = Math.Abs(nv) <= 1.000001 ? nv * 100 : nv;
            tb.Text = pv.ToString("0.##", CultureInfo.InvariantCulture);
        };
        Hook(tb, RaiseChanged);
    }

    private static bool AllowNumeric(TextBox tb, string input, bool integerOnly)
    {
        string candidate = tb.Text.Remove(tb.SelectionStart, tb.SelectionLength)
                                  .Insert(tb.SelectionStart, input);
        if (candidate is "-" or "." or "-.") return true;
        if (!double.TryParse(candidate, NumberStyles.Any, CultureInfo.InvariantCulture, out double v)) return false;
        if (integerOnly && candidate.Contains('.')) return false;
        _ = v;
        return true;
    }

    private void BuildDate(FieldEditor fe, FieldSchema schema, object? value)
    {
        ValueFormatter.TryToDateTime(value, out DateTime dt);
        bool hasValue = value != null && dt != default;
        bool withTime = schema.Type == FieldType.DateTime;

        var picker = new MultiTableDatePicker
        {
            SelectedDate = hasValue ? dt.Date : null,
            Padding = new Thickness(6, 4, 6, 4),
            SelectedDateFormat = DatePickerFormat.Short,
            VerticalContentAlignment = VerticalAlignment.Center
        };

        if (!withTime)
        {
            fe.Element = picker;
            fe.Read = () => picker.SelectedDate;
            fe.Write = v =>
            {
                picker.SelectedDate = ValueFormatter.TryToDateTime(v, out DateTime nd) ? nd.Date : null;
            };
            picker.SelectedDateChanged += (_, _) => RaiseChanged();
            return;
        }

        var timeBox = new TextBox
        {
            Width = 66,
            Padding = new Thickness(6, 6, 6, 6),
            VerticalContentAlignment = VerticalAlignment.Center,
            Text = hasValue ? dt.ToString("HH:mm") : string.Empty
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(timeBox, "HH:mm");

        var dock = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(timeBox, Dock.Right);
        timeBox.Margin = new Thickness(6, 0, 0, 0);
        dock.Children.Add(timeBox);
        dock.Children.Add(picker);

        fe.Element = dock;
        fe.Read = () =>
        {
            if (picker.SelectedDate == null) return null;
            var d = picker.SelectedDate.Value.Date;
            if (TimeSpan.TryParse(timeBox.Text, out TimeSpan ts)) d = d.Add(ts);
            return d;
        };
        fe.Write = v =>
        {
            if (!ValueFormatter.TryToDateTime(v, out DateTime nd)) { picker.SelectedDate = null; timeBox.Text = string.Empty; return; }
            picker.SelectedDate = nd.Date;
            timeBox.Text = nd.ToString("HH:mm");
        };
        picker.SelectedDateChanged += (_, _) => RaiseChanged();
        Hook(timeBox, RaiseChanged);
    }

    private void BuildCombo(FieldEditor fe, IEnumerable<string> options, object? value, bool editable, string placeholder)
    {
        var list = options?.ToList() ?? new List<string>();
        string current = ValueFormatter.ToDisplayText(value).Trim();
        if (current.Length > 0 && !list.Contains(current)) list.Insert(0, current);

        var combo = new ComboBox
        {
            ItemsSource = list,
            IsEditable = editable,
            Padding = new Thickness(8, 6, 8, 6),
            VerticalContentAlignment = VerticalAlignment.Center
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(combo, placeholder);

        if (current.Length > 0)
        {
            combo.SelectedItem = current;
            if (editable) combo.Text = current;
        }

        fe.Element = combo;
        fe.Read = () =>
        {
            string? v = editable ? combo.Text : combo.SelectedItem as string;
            return string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        };
        fe.Write = v =>
        {
            string s = ValueFormatter.ToDisplayText(v);
            if (editable) combo.Text = s; else combo.SelectedItem = s;
        };
        combo.SelectionChanged += (_, _) => RaiseChanged();
        if (editable) combo.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler((_, _) => RaiseChanged()));
    }

    private void BuildCheckbox(FieldEditor fe, object? value)
    {
        bool state = value switch
        {
            bool b => b,
            null => false,
            _ => ValueFormatter.ToDisplayText(value).Trim() is "是" or "true" or "TRUE" or "1" or "√" or "Y"
        };

        var cb = new CheckBox
        {
            IsChecked = state,
            Content = state ? "是" : "否",
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(2, 8, 0, 8)
        };
        cb.Checked += (_, _) => { cb.Content = "是"; RaiseChanged(); };
        cb.Unchecked += (_, _) => { cb.Content = "否"; RaiseChanged(); };

        fe.Element = cb;
        fe.Read = () => cb.IsChecked == true;
        fe.Write = v => cb.IsChecked = v is bool bb && bb;
    }

    private void BuildImagePath(FieldEditor fe, object? value)
    {
        var tb = new TextBox
        {
            Text = ValueFormatter.ToDisplayText(value),
            Padding = new Thickness(8, 6, 8, 6),
            VerticalContentAlignment = VerticalAlignment.Center
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "图片文件路径");

        var browse = new Button
        {
            Content = "浏览",
            Width = 56,
            Margin = new Thickness(6, 0, 0, 0),
            Style = TryFindResource("ButtonDefault") as Style
        };
        browse.Click += (_, _) =>
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*"
            };
            if (dlg.ShowDialog() == true) tb.Text = dlg.FileName;
        };

        var dock = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(browse, Dock.Right);
        dock.Children.Add(browse);
        dock.Children.Add(tb);

        fe.Element = dock;
        fe.Read = () => string.IsNullOrWhiteSpace(tb.Text) ? null : tb.Text.Trim();
        fe.Write = v => tb.Text = ValueFormatter.ToDisplayText(v);
        Hook(tb, RaiseChanged);
    }

    // ─────────────────────────────────────────────────────────────
    // 数量 × 单价 = 金额 联动
    // ─────────────────────────────────────────────────────────────

    private void OnNumericChanged(FieldEditor changed)
    {
        if (_suppressLink || _config == null) return;

        var link = _config.FindNumericLink(changed.Name);
        if (link == null) return;

        var qty = _editors.Find(e => e.Name == link.QuantityField);
        var price = _editors.Find(e => e.Name == link.UnitPriceField);
        var amount = _editors.Find(e => e.Name == link.AmountField);
        if (qty == null || price == null || amount == null) return;

        bool HasNum(FieldEditor e, out double v) => ValueFormatter.TryToDouble(e.Read(), out v);

        _suppressLink = true;
        try
        {
            if (changed.Name == link.AmountField)
            {
                // 改金额 → 反推单价
                if (HasNum(amount, out double a) && HasNum(qty, out double q) && Math.Abs(q) > 1e-9)
                    price.Write(Math.Round(a / q, link.UnitPriceDecimals));
            }
            else
            {
                // 改数量或单价 → 重算金额
                if (HasNum(qty, out double q) && HasNum(price, out double p))
                    amount.Write(Math.Round(q * p, link.AmountDecimals));
            }
        }
        finally { _suppressLink = false; }
    }

    private void RaiseChanged()
    {
        if (_suppressLink) return;
        ValueChanged?.Invoke(this, EventArgs.Empty);
    }

    // ─────────────────────────────────────────────────────────────
    // 取值
    // ─────────────────────────────────────────────────────────────

    /// <summary>获取所有字段的当前值</summary>
    public Dictionary<string, object?> GetValues()
    {
        var dict = new Dictionary<string, object?>(StringComparer.Ordinal);
        foreach (var e in _editors) dict[e.Name] = e.Read();
        return dict;
    }

    /// <summary>只返回与原值不同的字段</summary>
    public Dictionary<string, object?> GetChanges()
    {
        var dict = new Dictionary<string, object?>(StringComparer.Ordinal);
        foreach (var e in _editors)
        {
            object? now = e.Read();
            if (!SameValue(e.Original, now)) dict[e.Name] = now;
        }
        return dict;
    }

    /// <summary>按字段覆盖配置中的校验规则校验当前值，返回字段名→错误信息</summary>
    public Dictionary<string, string> Validate()
    {
        var errors = new Dictionary<string, string>(StringComparer.Ordinal);
        if (_config == null) return errors;

        foreach (var e in _editors)
        {
            var ov = _config.GetFieldOverride(e.Name);
            if (ov == null) continue;

            var err = FieldValidator.Validate(ov, e.Read());
            if (err != null) errors[e.Name] = err;

            // 内置字段名强校验：年龄 / 电话
            if (err == null && e.Name.Contains("年龄", StringComparison.OrdinalIgnoreCase))
            {
                var ageErr = FieldValidator.ValidateAge(e.Read());
                if (ageErr != null) errors[e.Name] = ageErr;
            }
            if (err == null && (e.Name.Contains("手机", StringComparison.OrdinalIgnoreCase) ||
                                e.Name.Contains("电话", StringComparison.OrdinalIgnoreCase)))
            {
                var phoneErr = FieldValidator.ValidatePhone(e.Read());
                if (phoneErr != null) errors[e.Name] = phoneErr;
            }
        }
        return errors;
    }

    public bool HasChanges => GetChanges().Count > 0;

    /// <summary>保存成功后把当前值作为新的基准</summary>
    public void AcceptChanges()
    {
        foreach (var e in _editors) e.Original = e.Read();
    }

    private static bool SameValue(object? a, object? b)
    {
        if (a == null && b == null) return true;
        if (a == null || b == null)
        {
            string other = ValueFormatter.ToDisplayText(a ?? b);
            return other.Length == 0;
        }

        if (a is DateTime da && b is DateTime db) return da == db;

        if (ValueFormatter.TryToDouble(a, out double na) && ValueFormatter.TryToDouble(b, out double nb))
            return Math.Abs(na - nb) < 1e-9;

        return string.Equals(ValueFormatter.ToDisplayText(a), ValueFormatter.ToDisplayText(b), StringComparison.Ordinal);
    }
}
