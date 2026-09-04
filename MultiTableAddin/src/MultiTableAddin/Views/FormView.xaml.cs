using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MultiTableAddin.Core;
using MultiTableAddin.Views.Controls;
using UserControl = System.Windows.Controls.UserControl;
using TextBox = System.Windows.Controls.TextBox;
using StackPanel = System.Windows.Controls.StackPanel;
using FontFamily = System.Windows.Media.FontFamily;

namespace MultiTableAddin.Views;

public partial class FormView : UserControl, ITableView, IConfigAware
{
    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;
    private ViewConfigFile? _configFile;
    private int _currentIndex;
    private readonly Dictionary<string, System.Windows.FrameworkElement> _inputControls = new();

    public FormView()
    {
        InitializeComponent();
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;
        _currentIndex = 0;
        ShowCurrentRecord();
    }

    public void SetConfigFile(ViewConfigFile configFile)
    {
        _configFile = configFile;
        if (_dataTable != null) ShowCurrentRecord();
    }

    private void ShowCurrentRecord()
    {
        FormPanel.Children.Clear();
        _inputControls.Clear();

        if (_viewData == null || _dataTable == null) return;

        var allRows = _viewData.Groups.SelectMany(g => g.Rows).ToList();
        if (allRows.Count == 0)
        {
            FormPanel.Children.Add(new TextBlock { Text = "没有数据", FontSize = 14 });
            return;
        }

        if (_currentIndex >= allRows.Count) _currentIndex = allRows.Count - 1;
        if (_currentIndex < 0) _currentIndex = 0;

        var row = allRows[_currentIndex];
        RecordInfo.Text = $"{_currentIndex + 1} / {allRows.Count}";

        var visibleFields = _viewConfig?.VisibleFields?.Count > 0
            ? _viewConfig.VisibleFields
            : _dataTable.Fields.ConvertAll(f => f.Name);

        // 根据实际可见字段名计算标签列宽度，避免固定 110 导致短字段名右侧大片空白
        double labelWidth = 80;
        foreach (var fieldName in visibleFields)
        {
            if (_dataTable.Fields.Find(f => f.Name == fieldName) == null) continue;
            labelWidth = Math.Max(labelWidth, MeasureText(fieldName, 13) + 24);
        }
        labelWidth = Math.Min(labelWidth, 220);

        foreach (var fieldName in visibleFields)
        {
            var field = GetEffectiveField(fieldName);
            if (field == null) continue;

            // 横向布局：标签在左，控件在右
            var rowPanel = new Grid { Margin = new Thickness(0, 6, 0, 6) };
            rowPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(labelWidth) });
            rowPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // 字段标签
            var label = new TextBlock
            {
                Text = fieldName,
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = TryFindResource("PrimaryTextBrush") as Brush,
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = fieldName
            };
            Grid.SetColumn(label, 0);
            rowPanel.Children.Add(label);

            // 根据字段类型创建输入控件
            var value = row.GetValue(fieldName);
            System.Windows.FrameworkElement inputControl = field.Type switch
            {
                FieldType.Select => CreateSelectControl(field, value),
                FieldType.Date => CreateDatePicker(value),
                FieldType.Quarter => CreateQuarterControl(value),
                FieldType.Currency => CreateCurrencyInput(value),
                FieldType.Percentage => CreatePercentageInput(value),
                FieldType.Number => CreateNumberInput(value),
                FieldType.Email => CreateEmailInput(value),
                FieldType.Phone => CreatePhoneInput(value),
                _ => CreateTextInput(value)
            };

            inputControl.VerticalAlignment = VerticalAlignment.Center;
            inputControl.HorizontalAlignment = HorizontalAlignment.Left;
            Grid.SetColumn(inputControl, 1);
            _inputControls[fieldName] = inputControl;
            rowPanel.Children.Add(inputControl);

            FormPanel.Children.Add(rowPanel);
        }
    }

    /// <summary>校验当前记录，返回字段名→错误信息</summary>
    private Dictionary<string, string> ValidateCurrentRecord()
    {
        var errors = new Dictionary<string, string>(StringComparer.Ordinal);
        if (_configFile == null) return errors;

        foreach (var kv in _inputControls)
        {
            string name = kv.Key;
            var control = kv.Value;
            object? value = control switch
            {
                TextBox tb => (object?)tb.Text,
                DatePicker dp => (object?)dp.SelectedDate,
                ComboBox cb => (object?)(cb.IsEditable ? cb.Text : cb.SelectedItem?.ToString()),
                System.Windows.Controls.Border bd when bd.Tag is TextBox tb2 => (object?)tb2.Text,
                _ => null
            };

            var ov = _configFile.GetFieldOverride(name);
            if (ov != null)
            {
                var err = FieldValidator.Validate(ov, value);
                if (err != null) errors[name] = err;
            }

            // 内置字段名强校验
            if (!errors.ContainsKey(name) && name.Contains("年龄", StringComparison.OrdinalIgnoreCase))
            {
                var ageErr = FieldValidator.ValidateAge(value);
                if (ageErr != null) errors[name] = ageErr;
            }
            if (!errors.ContainsKey(name) && (name.Contains("手机", StringComparison.OrdinalIgnoreCase) ||
                                              name.Contains("电话", StringComparison.OrdinalIgnoreCase)))
            {
                var phoneErr = FieldValidator.ValidatePhone(value);
                if (phoneErr != null) errors[name] = phoneErr;
            }
        }
        return errors;
    }

    /// <summary>获取字段的有效类型（考虑用户覆盖配置）</summary>
    private FieldSchema? GetEffectiveField(string fieldName)
    {
        var field = _dataTable?.Fields.Find(f => f.Name == fieldName);
        if (field == null) return null;

        if (_configFile != null)
        {
            var override_ = _configFile.FieldOverrides?.Find(f => f.Name == fieldName);
            if (override_ != null)
            {
                return new FieldSchema
                {
                    Name = field.Name,
                    Type = override_.Type,
                    Options = override_.Options.Count > 0 ? override_.Options : field.Options
                };
            }
        }
        return field;
    }

    private System.Windows.Controls.Control CreateTextInput(object? value)
    {
        var tb = new TextBox
        {
            Text = value?.ToString() ?? "",
            Padding = new Thickness(8, 6, 8, 6),
            MinWidth = 240,
            MaxWidth = 340,
            HorizontalAlignment = HorizontalAlignment.Left
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "请输入...");
        return tb;
    }

    private System.Windows.Controls.Control CreateNumberInput(object? value)
    {
        var tb = new TextBox
        {
            Text = value?.ToString() ?? "",
            Padding = new Thickness(8, 6, 8, 6),
            MinWidth = 140,
            MaxWidth = 190,
            HorizontalAlignment = HorizontalAlignment.Left
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "请输入数字");
        tb.PreviewTextInput += (s, e) =>
        {
            // 只允许数字、小数点、负号
            e.Handled = !IsValidNumberInput(e.Text, tb.Text, tb.CaretIndex);
        };
        return tb;
    }

    private static bool IsValidNumberInput(string newText, string currentText, int caretIndex)
    {
        string combined = currentText.Insert(caretIndex, newText);
        return double.TryParse(combined, out _);
    }

    private static double MeasureText(string text, double fontSize)
    {
        if (string.IsNullOrEmpty(text)) return fontSize;
        var tf = new Typeface(new FontFamily("Microsoft YaHei UI, Microsoft YaHei, Segoe UI"),
            FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        var ft = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            tf, fontSize, Brushes.Black, 1.0);
        return ft.Width;
    }

    private System.Windows.Controls.Border CreateCurrencyInput(object? value)
    {
        // 把货币符号放到输入框内部左侧，保持标签右边缘与输入框左边缘对齐
        var grid = new Grid { MinWidth = 110, MaxWidth = 150, HorizontalAlignment = HorizontalAlignment.Left };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(26) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var prefix = new TextBlock
        {
            Text = "¥",
            FontSize = 14,
            FontWeight = FontWeights.Medium,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = TryFindResource("PrimaryBrush") as Brush
        };
        Grid.SetColumn(prefix, 0);

        var tb = new TextBox
        {
            Text = value?.ToString() ?? "",
            Padding = new Thickness(4, 6, 8, 6),
            MinWidth = 90,
            BorderThickness = new Thickness(0)
        };
        Grid.SetColumn(tb, 1);
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "金额");
        tb.PreviewTextInput += (s, e) =>
        {
            e.Handled = !IsValidNumberInput(e.Text, tb.Text, tb.CaretIndex);
        };

        grid.Children.Add(prefix);
        grid.Children.Add(tb);

        var border = new System.Windows.Controls.Border
        {
            BorderBrush = TryFindResource("BorderBrush") as Brush,
            Background = TryFindResource("RegionBrush") as Brush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Child = grid
        };
        // 用 Tag 存储实际的 TextBox 便于读取值
        border.Tag = tb;
        return border;
    }

    private System.Windows.Controls.Border CreatePercentageInput(object? value)
    {
        // 把百分号放到输入框内部右侧
        var grid = new Grid { MinWidth = 110, MaxWidth = 150, HorizontalAlignment = HorizontalAlignment.Left };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(26) });

        var tb = new TextBox
        {
            Text = value?.ToString() ?? "",
            Padding = new Thickness(8, 6, 4, 6),
            MinWidth = 90,
            BorderThickness = new Thickness(0)
        };
        Grid.SetColumn(tb, 0);
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "百分比");
        tb.PreviewTextInput += (s, e) =>
        {
            e.Handled = !IsValidNumberInput(e.Text, tb.Text, tb.CaretIndex);
        };

        var suffix = new TextBlock
        {
            Text = "%",
            FontSize = 14,
            FontWeight = FontWeights.Medium,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
        };
        Grid.SetColumn(suffix, 1);

        grid.Children.Add(tb);
        grid.Children.Add(suffix);

        var border = new System.Windows.Controls.Border
        {
            BorderBrush = TryFindResource("BorderBrush") as Brush,
            Background = TryFindResource("RegionBrush") as Brush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Child = grid
        };
        border.Tag = tb;
        return border;
    }

    private System.Windows.Controls.Control CreateEmailInput(object? value)
    {
        var tb = new TextBox
        {
            Text = value?.ToString() ?? "",
            Padding = new Thickness(8, 6, 8, 6),
            MinWidth = 240,
            MaxWidth = 340,
            HorizontalAlignment = HorizontalAlignment.Left
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "请输入邮箱地址");
        return tb;
    }

    private System.Windows.Controls.Control CreatePhoneInput(object? value)
    {
        var tb = new TextBox
        {
            Text = value?.ToString() ?? "",
            Padding = new Thickness(8, 6, 8, 6),
            MinWidth = 140,
            MaxWidth = 190,
            HorizontalAlignment = HorizontalAlignment.Left,
            MaxLength = 11
        };
        HandyControl.Controls.InfoElement.SetPlaceholder(tb, "请输入手机号码");
        tb.PreviewTextInput += (s, e) =>
        {
            // 只允许数字
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c)) { e.Handled = true; break; }
            }
        };
        return tb;
    }

    private System.Windows.Controls.Control CreateDatePicker(object? value)
    {
        var picker = new MultiTableDatePicker
        {
            SelectedDate = value is DateTime dt ? dt : null,
            Padding = new Thickness(8, 6, 8, 6),
            SelectedDateFormat = DatePickerFormat.Short,
            MinWidth = 220,
            MaxWidth = 300,
            HorizontalAlignment = HorizontalAlignment.Left
        };
        return picker;
    }

    private System.Windows.Controls.Control CreateSelectControl(FieldSchema field, object? value)
    {
        var combo = new ComboBox
        {
            Padding = new Thickness(8, 6, 8, 6),
            ItemsSource = field.Options,
            MinWidth = 220,
            MaxWidth = 300,
            HorizontalAlignment = HorizontalAlignment.Left
        };
        if (value != null)
        {
            combo.SelectedItem = value.ToString();
        }
        return combo;
    }

    private System.Windows.Controls.Control CreateQuarterControl(object? value)
    {
        var combo = new ComboBox
        {
            Padding = new Thickness(8, 6, 8, 6),
            ItemsSource = FieldSchema.QuarterOptions,
            MinWidth = 220,
            MaxWidth = 300,
            HorizontalAlignment = HorizontalAlignment.Left
        };
        if (value != null)
        {
            combo.SelectedItem = value.ToString();
        }
        return combo;
    }

    private void OnPrevClick(object sender, RoutedEventArgs e)
    {
        if (_currentIndex > 0)
        {
            _currentIndex--;
            ShowCurrentRecord();
        }
    }

    private void OnNextClick(object sender, RoutedEventArgs e)
    {
        if (_viewData == null) return;
        var count = _viewData.Groups.Sum(g => g.Rows.Count);
        if (_currentIndex < count - 1)
        {
            _currentIndex++;
            ShowCurrentRecord();
        }
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        if (_dataTable == null || _excelAdapter == null || _viewData == null) return;

        var allRows = _viewData.Groups.SelectMany(g => g.Rows).ToList();
        if (_currentIndex < 0 || _currentIndex >= allRows.Count) return;

        // 校验
        var errors = ValidateCurrentRecord();
        if (errors.Count > 0)
        {
            HandyControl.Controls.Growl.WarningGlobal("校验未通过：\n" + string.Join("\n", errors.Select(kv => $"• {kv.Key}: {kv.Value}")));
            return;
        }

        var row = allRows[_currentIndex];

        foreach (var kv in _inputControls)
        {
            string fieldName = kv.Key;
            var control = kv.Value;
            object? newValue = control switch
            {
                TextBox tb => (object?)tb.Text,
                DatePicker dp => (object?)dp.SelectedDate,
                ComboBox cb => (object?)cb.SelectedItem,
                System.Windows.Controls.Border bd when bd.Tag is TextBox tb2 => (object?)tb2.Text,
                _ => null
            };

            if (newValue != null && !Equals(row.GetValue(fieldName), newValue))
            {
                row.SetValue(fieldName, newValue);
                _excelAdapter.UpdateCell(_dataTable.SheetName, _dataTable.TableName, row.RowIndex, fieldName, newValue);
                _dataTable.IsDirty = true;
            }
        }

        HandyControl.Controls.Growl.SuccessGlobal("修改已保存到 Excel。");
    }
}
