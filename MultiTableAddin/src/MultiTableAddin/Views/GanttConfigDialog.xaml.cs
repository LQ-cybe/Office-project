using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MultiTableAddin.Core;
using CheckBox = System.Windows.Controls.CheckBox;

namespace MultiTableAddin.Views;

public partial class GanttConfigDialog : Window
{
    private readonly DataTableModel _dataTable;
    public GanttConfig Config { get; private set; }

    private readonly List<string> _allFields;

    public GanttConfigDialog(DataTableModel dataTable, GanttConfig config)
    {
        _dataTable = dataTable;
        Config = config.Clone();
        _allFields = dataTable.Fields.ConvertAll(f => f.Name);
        InitializeComponent();
        LoadData();

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        };
    }

    private void LoadData()
    {
        FillCombo(StartFieldCombo, _allFields, Config.StartField);
        FillCombo(EndFieldCombo, _allFields, Config.EndField);
        FillCombo(LabelFieldCombo, _allFields, Config.LabelField);

        var groupFields = new List<string> { "(不分组)" };
        groupFields.AddRange(_allFields);
        FillCombo(GroupFieldCombo, groupFields, string.IsNullOrEmpty(Config.GroupField) ? "(不分组)" : Config.GroupField);

        // 显示字段：使用独立 CheckBox（直接点选即可，无需先选中），行距正常
        DisplayFieldsPanel.Children.Clear();
        foreach (var f in _allFields)
        {
            var cb = new CheckBox
            {
                Content = f,
                IsChecked = Config.DisplayFields.Contains(f),
                Margin = new Thickness(0, 3, 16, 3),
                VerticalAlignment = VerticalAlignment.Center
            };
            DisplayFieldsPanel.Children.Add(cb);
        }
    }

    private static void FillCombo(ComboBox combo, List<string> items, string selected)
    {
        combo.Items.Clear();
        foreach (var item in items)
        {
            combo.Items.Add(item);
        }
        combo.SelectedItem = items.Contains(selected) ? selected : items.FirstOrDefault();
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        Config.StartField = StartFieldCombo.SelectedItem?.ToString() ?? "";
        Config.EndField = EndFieldCombo.SelectedItem?.ToString() ?? "";
        Config.LabelField = LabelFieldCombo.SelectedItem?.ToString() ?? "";
        Config.GroupField = GroupFieldCombo.SelectedItem?.ToString() == "(不分组)" ? "" : GroupFieldCombo.SelectedItem?.ToString() ?? "";
        Config.DisplayFields = DisplayFieldsPanel.Children.OfType<CheckBox>()
            .Where(c => c.IsChecked == true)
            .Select(c => c.Content?.ToString() ?? "")
            .Where(s => !string.IsNullOrEmpty(s))
            .ToList();

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}

public static class GanttConfigExtensions
{
    public static GanttConfig Clone(this GanttConfig cfg) => new()
    {
        StartField = cfg.StartField,
        EndField = cfg.EndField,
        LabelField = cfg.LabelField,
        GroupField = cfg.GroupField,
        ProgressField = cfg.ProgressField,
        TimeDimension = cfg.TimeDimension,
        DisplayFields = new List<string>(cfg.DisplayFields)
    };
}
