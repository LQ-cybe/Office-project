using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using MultiTableAddin.Core;

namespace MultiTableAddin.Views;

/// <summary>仪表盘 KPI 指标卡设置窗口</summary>
public partial class StatCardDialog : Window
{
    private const string NoneText = "（记录数，不指定字段）";

    private static readonly (string Key, string Label)[] FormatOptions =
    {
        ("auto", "自动（万 / 亿）"),
        ("int",  "整数"),
        ("money","金额 ¥"),
        ("percent","百分比")
    };

    private static readonly (string? Hex, string Label)[] ColorOptions =
    {
        (null,      "自动（按顺序取色）"),
        ("#4E7CF6", "蓝"),
        ("#27AE60", "绿"),
        ("#EB5757", "红"),
        ("#F2994A", "橙"),
        ("#9B51E0", "紫"),
        ("#00B8A9", "青"),
        ("#F2C94C", "黄")
    };

    private readonly DataTableModel _table;

    public StatCardConfig Result { get; }

    public StatCardDialog(DataTableModel table, ViewConfigFile? config, StatCardConfig source)
    {
        InitializeComponent();

        _table = table;
        Result = new StatCardConfig
        {
            Id = string.IsNullOrEmpty(source.Id) ? ViewConfig.NewId("stat") : source.Id,
            Title = source.Title,
            Field = source.Field,
            Aggregation = source.Aggregation,
            Filter = source.Filter,
            Format = source.Format,
            Color = source.Color
        };

        TitleBox.Text = Result.Title;
        FilterBox.Text = Result.Filter;

        // 字段：数值型排前面，方便选择
        FieldCombo.Items.Add(NoneText);
        var numeric = _table.Fields
            .Where(f => FieldTypeHelper.IsNumeric(config?.GetEffectiveFieldType(f.Name) ?? f.Type))
            .Select(f => f.Name).ToList();
        foreach (var n in numeric) FieldCombo.Items.Add(n);
        foreach (var f in _table.Fields.Where(f => !numeric.Contains(f.Name)))
            FieldCombo.Items.Add(f.Name);

        FieldCombo.SelectedItem = string.IsNullOrEmpty(Result.Field) || !FieldCombo.Items.Contains(Result.Field)
            ? NoneText
            : Result.Field;

        foreach (var kv in AggregateModeHelper.AllLabels) AggCombo.Items.Add(kv);
        AggCombo.SelectedValue = Result.Aggregation;
        if (AggCombo.SelectedIndex < 0) AggCombo.SelectedIndex = 0;

        foreach (var f in FormatOptions) FormatCombo.Items.Add(f.Label);
        int fi = Array.FindIndex(FormatOptions, f => f.Key == (Result.Format ?? "auto"));
        FormatCombo.SelectedIndex = fi < 0 ? 0 : fi;

        foreach (var c in ColorOptions) ColorCombo.Items.Add(c.Label);
        int ci = Array.FindIndex(ColorOptions, c =>
            string.Equals(c.Hex, Result.Color, StringComparison.OrdinalIgnoreCase));
        ColorCombo.SelectedIndex = ci < 0 ? 0 : ci;

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        };
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        string field = FieldCombo.SelectedItem as string ?? string.Empty;
        Result.Field = field == NoneText ? string.Empty : field;

        if (AggCombo.SelectedValue is AggregateMode mode) Result.Aggregation = mode;

        Result.Title = string.IsNullOrWhiteSpace(TitleBox.Text)
            ? (string.IsNullOrEmpty(Result.Field)
                ? AggregateModeHelper.GetLabel(Result.Aggregation)
                : Result.Field + " " + AggregateModeHelper.GetLabel(Result.Aggregation))
            : TitleBox.Text.Trim();

        int fi = FormatCombo.SelectedIndex;
        Result.Format = fi >= 0 && fi < FormatOptions.Length ? FormatOptions[fi].Key : "auto";

        int ci = ColorCombo.SelectedIndex;
        Result.Color = ci > 0 && ci < ColorOptions.Length ? ColorOptions[ci].Hex : null;

        Result.Filter = FilterBox.Text?.Trim() ?? string.Empty;

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
