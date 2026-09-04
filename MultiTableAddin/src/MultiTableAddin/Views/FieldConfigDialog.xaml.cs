using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MultiTableAddin.Core;
using TextBox = System.Windows.Controls.TextBox;
using CheckBox = System.Windows.Controls.CheckBox;

namespace MultiTableAddin.Views;

/// <summary>
/// 字段配置对话框（对应需求：字段类型/选项手动可配 + 跨工作簿规则库 + 数量×单价=金额联动）。
/// 三个页签：① 字段类型（逐字段覆盖）② 字段识别规则库（增删改、持久化）③ 数值联动。
/// </summary>
public partial class FieldConfigDialog : Window
{
    private readonly DataTableModel _table;
    private readonly ViewConfigFile _config;
    private FieldRuleLibrary? _rules;
    private bool _changed;
    private bool _rulesDirty;

    /// <summary>用户在对话框中是否做了任何修改（主视图据此刷新）</summary>
    public bool ConfigChanged => _changed || _rulesDirty;

    public FieldConfigDialog(DataTableModel dataTable, ViewConfigFile configFile)
    {
        InitializeComponent();
        _table = dataTable;
        _config = configFile;
        BuildFieldRows();
        LoadRules();
        LoadLinks();
        LoadValidations();

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        };
    }

    // ─────────────────────────────────────────────────────────────
    // 页签一：字段类型
    // ─────────────────────────────────────────────────────────────

    private void BuildFieldRows()
    {
        FieldRows.Children.Clear();
        foreach (var field in _table.Fields)
        {
            var eff = _config.GetEffectiveField(field.Name);

            var row = new Grid { Margin = new Thickness(0, 3, 0, 3) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var name = new TextBlock
            {
                Text = field.Name,
                FontSize = 12,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = field.Name,
                Foreground = Res("PrimaryTextBrush")
            };
            Grid.SetColumn(name, 0);

            var typeCombo = new ComboBox
            {
                Margin = new Thickness(6, 0, 6, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                ItemsSource = FieldTypeHelper.AllLabels.ToList(),
                DisplayMemberPath = "Value",
                SelectedValuePath = "Key",
                SelectedValue = eff.Type
            };

            var optBox = new TextBox
            {
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(6, 4, 6, 4),
                Text = string.Join("; ", eff.Options),
                IsEnabled = eff.Type == FieldType.Select
            };
            HandyControl.Controls.InfoElement.SetPlaceholder(optBox, "选项用 ; 分隔");

            typeCombo.SelectionChanged += (_, _) =>
            {
                var t = (FieldType)(typeCombo.SelectedValue ?? FieldType.Text);
                optBox.IsEnabled = t == FieldType.Select;
                if (t != FieldType.Select) optBox.Text = string.Empty;
                ApplyFieldOverride(field, t, optBox.Text);
            };
            optBox.LostFocus += (_, _) =>
                ApplyFieldOverride(field, (FieldType)typeCombo.SelectedValue!, optBox.Text);

            Grid.SetColumn(typeCombo, 1);
            Grid.SetColumn(optBox, 2);
            row.Children.Add(name);
            row.Children.Add(typeCombo);
            row.Children.Add(optBox);
            FieldRows.Children.Add(row);
        }
    }

    private void ApplyFieldOverride(FieldSchema field, FieldType type, string optionsText)
    {
        var opts = new List<string>();
        if (type == FieldType.Select && !string.IsNullOrWhiteSpace(optionsText))
            opts = optionsText.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries)
                              .Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        _config.SetFieldOverride(field.Name, type, opts);
        _changed = true;
    }

    // ─────────────────────────────────────────────────────────────
    // 页签二：字段识别规则库
    // ─────────────────────────────────────────────────────────────

    private class RuleViewItem
    {
        public FieldRule Rule { get; }
        public string Remark => Rule.Remark;
        public string Keywords => Rule.Keywords;
        public string MatchModeLabel => RuleMatchModeHelper.GetLabel(Rule.MatchMode);
        public string TypeLabel => FieldTypeHelper.GetLabel(Rule.Type);
        public string EnabledLabel => Rule.Enabled ? "是" : "否";
        public RuleViewItem(FieldRule r) => Rule = r;
    }

    private void LoadRules()
    {
        _rules = FieldRuleLibrary.Current;
        RuleList.ItemsSource = _rules.Rules
            .OrderByDescending(r => r.Priority)
            .Select(r => new RuleViewItem(r))
            .ToList();
    }

    private RuleViewItem? SelectedRule()
    {
        if (RuleList.SelectedItem is RuleViewItem item) return item;
        HandyControl.Controls.Growl.WarningGlobal("请先选择一条规则");
        return null;
    }

    private void OnAddRule(object sender, RoutedEventArgs e)
    {
        var r = new FieldRule
        {
            Id = "rule_" + Guid.NewGuid().ToString("N").Substring(0, 8),
            Enabled = true,
            Priority = 100
        };
        var dlg = new RuleEditWindow(r) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _rules!.Rules.Add(r);
            _rules.Save();
            _rulesDirty = true;
            LoadRules();
        }
    }

    private void OnEditRule(object sender, RoutedEventArgs e)
    {
        var item = SelectedRule();
        if (item == null) return;
        var dlg = new RuleEditWindow(item.Rule) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _rules!.Save();
            _rulesDirty = true;
            LoadRules();
        }
    }

    private void OnRuleDoubleClick(object sender, MouseButtonEventArgs e) => OnEditRule(sender, e);

    private void OnDeleteRule(object sender, RoutedEventArgs e)
    {
        var item = SelectedRule();
        if (item == null) return;
        if (item.Rule.BuiltIn &&
            HandyControl.Controls.MessageBox.Show("内置规则建议保留（可改为停用）。确定删除？",
                "删除规则", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
            return;
        _rules!.Rules.Remove(item.Rule);
        _rules.Save();
        _rulesDirty = true;
        LoadRules();
    }

    private void OnResetRule(object sender, RoutedEventArgs e)
    {
        if (HandyControl.Controls.MessageBox.Show(
                "将用内置默认规则覆盖当前的规则库（用户新增的规则会保留，内置规则恢复默认）。继续？",
                "重置规则库", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
            return;
        _rules = FieldRuleLibrary.ResetToDefault();
        _rulesDirty = true;
        LoadRules();
    }

    // ─────────────────────────────────────────────────────────────
    // 页签三：数量 × 单价 = 金额 联动
    // ─────────────────────────────────────────────────────────────

    private void LoadLinks()
    {
        LinkList.Items.Clear();
        foreach (var link in _config.NumericLinks)
            LinkList.Items.Add(BuildLinkItem(link));
    }

    private ListBoxItem BuildLinkItem(NumericLinkConfig link)
    {
        var item = new ListBoxItem { Tag = link };
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(2, 4, 2, 4) };
        panel.Children.Add(new TextBlock
        {
            Text = $"数量「{link.QuantityField}」× 单价「{link.UnitPriceField}」= 金额「{link.AmountField}」" +
                   $"（金额{link.AmountDecimals}位 / 单价{link.UnitPriceDecimals}位）",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Res("PrimaryTextBrush")
        });
        item.Content = panel;
        return item;
    }

    private NumericLinkConfig? SelectedLink()
    {
        if (LinkList.SelectedItem is ListBoxItem item && item.Tag is NumericLinkConfig link) return link;
        HandyControl.Controls.Growl.WarningGlobal("请先选择一条联动");
        return null;
    }

    private void OnAddLink(object sender, RoutedEventArgs e)
    {
        var link = new NumericLinkConfig();
        var dlg = new LinkEditWindow(link, _table.FieldNames) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            if (!link.IsValid) { HandyControl.Controls.Growl.WarningGlobal("三个字段都需选择"); return; }
            _config.NumericLinks.Add(link);
            _changed = true;
            LoadLinks();
        }
    }

    private void OnEditLink(object sender, RoutedEventArgs e)
    {
        var link = SelectedLink();
        if (link == null) return;
        var dlg = new LinkEditWindow(link, _table.FieldNames) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _changed = true;
            LoadLinks();
        }
    }

    private void OnDeleteLink(object sender, RoutedEventArgs e)
    {
        var link = SelectedLink();
        if (link == null) return;
        _config.NumericLinks.Remove(link);
        _changed = true;
        LoadLinks();
    }

    // ─────────────────────────────────────────────────────────────
    // 确定 / 取消
    // ─────────────────────────────────────────────────────────────

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private Brush? Res(string key) => TryFindResource(key) as Brush;

    // ─────────────────────────────────────────────────────────────
    // 字段校验规则页签
    // ─────────────────────────────────────────────────────────────

    private void LoadValidations()
    {
        var items = new List<ValidationViewItem>();
        foreach (var field in _table.Fields)
        {
            var ov = _config.GetFieldOverride(field.Name);
            items.Add(new ValidationViewItem(field, ov));
        }
        ValidationList.ItemsSource = items;
    }

    private void OnValidationDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ValidationList.SelectedItem is not ValidationViewItem item) return;
        var field = _table.Fields.Find(f => f.Name == item.FieldName);
        if (field == null) return;

        var ov = _config.GetFieldOverride(field.Name);
        if (ov == null)
        {
            ov = new FieldOverride { Name = field.Name, Type = field.Type, UserDefined = false };
            _config.FieldOverrides.Add(ov);
        }

        var dlg = new ValidationEditWindow(ov);
        dlg.Owner = this;
        if (dlg.ShowDialog() == true)
        {
            _changed = true;
            LoadValidations();
        }
    }

    public class ValidationViewItem
    {
        public string FieldName { get; }
        public string RequiredText { get; }
        public string MinValueText { get; }
        public string MaxValueText { get; }
        public string MinLengthText { get; }
        public string MaxLengthText { get; }
        public string RegexText { get; }

        public ValidationViewItem(FieldSchema field, FieldOverride? ov)
        {
            FieldName = field.Name;
            RequiredText = ov?.Required == true ? "是" : "";
            MinValueText = ov?.MinValue?.ToString() ?? "";
            MaxValueText = ov?.MaxValue?.ToString() ?? "";
            MinLengthText = ov?.MinLength?.ToString() ?? "";
            MaxLengthText = ov?.MaxLength?.ToString() ?? "";
            RegexText = ov?.RegexPattern ?? "";
        }
    }

    // ─────────────────────────────────────────────────────────────
    // 规则编辑模态窗（纯代码构建，避免额外 XAML）
    // ─────────────────────────────────────────────────────────────

    private class RuleEditWindow : Window
    {
        private readonly FieldRule _rule;
        private readonly ComboBox _typeCombo;
        private readonly ComboBox _modeCombo;
        private readonly TextBox _remarkBox, _keywordsBox, _optionsBox, _priorityBox;
        private readonly CheckBox _enabledBox;

        public RuleEditWindow(FieldRule rule)
        {
            _rule = rule;
            Title = "编辑字段识别规则";
            Width = 430; Height = 340;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Background = (Brush)FindResource("RegionBrush")!;

            var dock = new DockPanel { Margin = new Thickness(14) };

            var bar = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            DockPanel.SetDock(bar, Dock.Bottom);
            var ok = new Button { Content = "确定", Width = 88, Style = (Style)FindResource("ButtonPrimary")! };
            ok.Click += (_, _) => { if (Commit()) { DialogResult = true; Close(); } };
            var cancel = new Button { Content = "取消", Width = 88, Margin = new Thickness(10, 0, 0, 0), Style = (Style)FindResource("ButtonDefault")! };
            cancel.Click += (_, _) => { DialogResult = false; Close(); };
            bar.Children.Add(ok); bar.Children.Add(cancel);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(88) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            int r = 0;
            void Row(string label, UIElement ctrl)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var tb = new TextBlock
                {
                    Text = label,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 0, 8, 0),
                    Foreground = (Brush)FindResource("SecondaryTextBrush")!
                };
                Grid.SetRow(tb, r); Grid.SetColumn(tb, 0);
                Grid.SetRow(ctrl, r); Grid.SetColumn(ctrl, 1);
                grid.Children.Add(tb); grid.Children.Add(ctrl);
                r++;
            }

            _remarkBox = Txt(_rule.Remark);
            _keywordsBox = Txt(_rule.Keywords);
            _modeCombo = new ComboBox { ItemsSource = RuleMatchModeHelper.AllLabels.ToList(), DisplayMemberPath = "Value", SelectedValuePath = "Key", SelectedValue = _rule.MatchMode, VerticalContentAlignment = VerticalAlignment.Center };
            _typeCombo = new ComboBox { ItemsSource = FieldTypeHelper.AllLabels.ToList(), DisplayMemberPath = "Value", SelectedValuePath = "Key", SelectedValue = _rule.Type, VerticalContentAlignment = VerticalAlignment.Center };
            _optionsBox = Txt(_rule.Options);
            _enabledBox = new CheckBox { Content = "启用", IsChecked = _rule.Enabled, VerticalAlignment = VerticalAlignment.Center };
            _priorityBox = Txt(_rule.Priority.ToString());

            Row("说明", _remarkBox);
            Row("关键词(;分隔)", _keywordsBox);
            Row("匹配方式", _modeCombo);
            Row("字段类型", _typeCombo);
            Row("选项(;分隔)", _optionsBox);
            Row("优先级", _priorityBox);
            Row("状态", _enabledBox);

            dock.Children.Add(bar);
            dock.Children.Add(grid);
            Content = dock;

            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    DialogResult = false;
                    Close();
                }
            };
        }

        private static TextBox Txt(string text) => new()
        {
            Text = text ?? "",
            Padding = new Thickness(6, 4, 6, 4),
            VerticalContentAlignment = VerticalAlignment.Center
        };

        private bool Commit()
        {
            _rule.Remark = _remarkBox.Text?.Trim() ?? "";
            _rule.Keywords = _keywordsBox.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(_rule.Keywords))
            {
                HandyControl.Controls.Growl.WarningGlobal("请填写关键词");
                return false;
            }
            if (_modeCombo.SelectedValue is RuleMatchMode m) _rule.MatchMode = m;
            if (_typeCombo.SelectedValue is FieldType t) _rule.Type = t;
            _rule.Options = _optionsBox.Text?.Trim() ?? "";
            _rule.Enabled = _enabledBox.IsChecked == true;
            if (int.TryParse(_priorityBox.Text, out int p)) _rule.Priority = p;
            return true;
        }
    }

    // ─────────────────────────────────────────────────────────────
    // 联动编辑模态窗（纯代码构建）
    // ─────────────────────────────────────────────────────────────

    private class LinkEditWindow : Window
    {
        private readonly NumericLinkConfig _link;
        private readonly ComboBox _q, _p, _a;
        private readonly TextBox _amtDec, _priceDec;

        public LinkEditWindow(NumericLinkConfig link, List<string> fields)
        {
            _link = link;
            Title = "编辑数值联动";
            Width = 430; Height = 280;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Background = (Brush)FindResource("RegionBrush")!;

            var dock = new DockPanel { Margin = new Thickness(14) };

            var bar = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            DockPanel.SetDock(bar, Dock.Bottom);
            var ok = new Button { Content = "确定", Width = 88, Style = (Style)FindResource("ButtonPrimary")! };
            ok.Click += (_, _) => { Commit(); DialogResult = true; Close(); };
            var cancel = new Button { Content = "取消", Width = 88, Margin = new Thickness(10, 0, 0, 0), Style = (Style)FindResource("ButtonDefault")! };
            cancel.Click += (_, _) => { DialogResult = false; Close(); };
            bar.Children.Add(ok); bar.Children.Add(cancel);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(96) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            int r = 0;
            void Row(string label, UIElement ctrl)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var tb = new TextBlock
                {
                    Text = label,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 0, 8, 0),
                    Foreground = (Brush)FindResource("SecondaryTextBrush")!
                };
                Grid.SetRow(tb, r); Grid.SetColumn(tb, 0);
                Grid.SetRow(ctrl, r); Grid.SetColumn(ctrl, 1);
                grid.Children.Add(tb); grid.Children.Add(ctrl);
                r++;
            }

            _q = FieldCombo(fields, _link.QuantityField);
            _p = FieldCombo(fields, _link.UnitPriceField);
            _a = FieldCombo(fields, _link.AmountField);
            _amtDec = new TextBox { Text = _link.AmountDecimals.ToString(), Padding = new Thickness(6, 4, 6, 4), VerticalContentAlignment = VerticalAlignment.Center };
            _priceDec = new TextBox { Text = _link.UnitPriceDecimals.ToString(), Padding = new Thickness(6, 4, 6, 4), VerticalContentAlignment = VerticalAlignment.Center };

            Row("数量字段", _q);
            Row("单价字段", _p);
            Row("金额字段", _a);
            Row("金额小数位", _amtDec);
            Row("单价小数位", _priceDec);

            dock.Children.Add(bar);
            dock.Children.Add(grid);
            Content = dock;

            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    DialogResult = false;
                    Close();
                }
            };
        }

        private static ComboBox FieldCombo(List<string> fields, string sel)
        {
            var cb = new ComboBox { ItemsSource = fields, VerticalContentAlignment = VerticalAlignment.Center };
            cb.SelectedItem = fields.Contains(sel) ? sel : (fields.Count > 0 ? fields[0] : null);
            return cb;
        }

        private void Commit()
        {
            _link.QuantityField = _q.SelectedItem as string ?? "";
            _link.UnitPriceField = _p.SelectedItem as string ?? "";
            _link.AmountField = _a.SelectedItem as string ?? "";
            if (int.TryParse(_amtDec.Text, out int d1)) _link.AmountDecimals = d1;
            if (int.TryParse(_priceDec.Text, out int d2)) _link.UnitPriceDecimals = d2;
        }
    }

    // ─────────────────────────────────────────────────────────────
    // 校验规则编辑模态窗
    // ─────────────────────────────────────────────────────────────

    private class ValidationEditWindow : Window
    {
        private readonly FieldOverride _ov;
        private readonly CheckBox _requiredBox;
        private readonly TextBox _minValueBox, _maxValueBox, _minLengthBox, _maxLengthBox, _regexBox, _errorBox;

        public ValidationEditWindow(FieldOverride ov)
        {
            _ov = ov;
            Title = $"校验规则 · {ov.Name}";
            Width = 460; Height = 360;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Background = (Brush)FindResource("RegionBrush")!;

            var dock = new DockPanel { Margin = new Thickness(14) };

            var bar = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            DockPanel.SetDock(bar, Dock.Bottom);
            var ok = new Button { Content = "确定", Width = 88, Style = (Style)FindResource("ButtonPrimary")! };
            ok.Click += (_, _) => { Commit(); DialogResult = true; Close(); };
            var cancel = new Button { Content = "取消", Width = 88, Margin = new Thickness(10, 0, 0, 0), Style = (Style)FindResource("ButtonDefault")! };
            cancel.Click += (_, _) => { DialogResult = false; Close(); };
            bar.Children.Add(ok); bar.Children.Add(cancel);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            int r = 0;
            void Row(string label, UIElement ctrl)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                var tb = new TextBlock
                {
                    Text = label,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 0, 8, 0),
                    Foreground = (Brush)FindResource("SecondaryTextBrush")!
                };
                Grid.SetRow(tb, r); Grid.SetColumn(tb, 0);
                Grid.SetRow(ctrl, r); Grid.SetColumn(ctrl, 1);
                grid.Children.Add(tb); grid.Children.Add(ctrl);
                r++;
            }

            _requiredBox = new CheckBox { Content = "必填", IsChecked = ov.Required, VerticalAlignment = VerticalAlignment.Center };
            _minValueBox = Txt(ov.MinValue?.ToString() ?? "");
            _maxValueBox = Txt(ov.MaxValue?.ToString() ?? "");
            _minLengthBox = Txt(ov.MinLength?.ToString() ?? "");
            _maxLengthBox = Txt(ov.MaxLength?.ToString() ?? "");
            _regexBox = Txt(ov.RegexPattern ?? "");
            _errorBox = Txt(ov.ErrorMessage ?? "");

            Row("必填", _requiredBox);
            Row("最小值", _minValueBox);
            Row("最大值", _maxValueBox);
            Row("最短长度", _minLengthBox);
            Row("最长长度", _maxLengthBox);
            Row("正则表达式", _regexBox);
            Row("错误提示", _errorBox);

            dock.Children.Add(bar);
            dock.Children.Add(grid);
            Content = dock;

            PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    DialogResult = false;
                    Close();
                }
            };
        }

        private static TextBox Txt(string text) => new()
        {
            Text = text ?? "",
            Padding = new Thickness(6, 4, 6, 4),
            VerticalContentAlignment = VerticalAlignment.Center
        };

        private void Commit()
        {
            _ov.Required = _requiredBox.IsChecked == true;
            _ov.MinValue = string.IsNullOrWhiteSpace(_minValueBox.Text) ? null
                : (double.TryParse(_minValueBox.Text, out double minD) ? minD : null);
            _ov.MaxValue = string.IsNullOrWhiteSpace(_maxValueBox.Text) ? null
                : (double.TryParse(_maxValueBox.Text, out double maxD) ? maxD : null);
            _ov.MinLength = string.IsNullOrWhiteSpace(_minLengthBox.Text) ? null
                : (int.TryParse(_minLengthBox.Text, out int minI) ? minI : null);
            _ov.MaxLength = string.IsNullOrWhiteSpace(_maxLengthBox.Text) ? null
                : (int.TryParse(_maxLengthBox.Text, out int maxI) ? maxI : null);
            _ov.RegexPattern = _regexBox.Text?.Trim() ?? string.Empty;
            _ov.ErrorMessage = _errorBox.Text?.Trim() ?? string.Empty;
        }
    }
}
