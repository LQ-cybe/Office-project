using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MultiTableAddin.Core;
using ListBox = System.Windows.Controls.ListBox;
using SelectionMode = System.Windows.Controls.SelectionMode;

namespace MultiTableAddin.Views;

/// <summary>新建视图时选择视图类型的对话框</summary>
public class AddViewDialog : Window
{
    public ViewType SelectedType { get; private set; } = ViewType.Table;

    private static readonly (ViewType type, string label)[] Options =
    {
        (ViewType.Table, "表格视图"),
        (ViewType.Form, "表单视图"),
        (ViewType.Kanban, "看板视图"),
        (ViewType.Gallery, "画册视图"),
        (ViewType.Calendar, "日历视图"),
        (ViewType.Gantt, "甘特视图"),
        (ViewType.Dashboard, "仪表盘"),
        (ViewType.Chart, "统计图表")
    };

    public AddViewDialog()
    {
        Title = "新建视图";
        Width = 300;
        Height = 420;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = TryFindResource("RegionBrush") as Brush ?? Brushes.White;

        var dock = new DockPanel { Margin = new Thickness(14) };

        var bar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        DockPanel.SetDock(bar, Dock.Bottom);

        var ok = new Button { Content = "确定", Width = 80, Style = (Style)FindResource("ButtonPrimary")! };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        var cancel = new Button
        {
            Content = "取消",
            Width = 80,
            Margin = new Thickness(10, 0, 0, 0),
            Style = (Style)FindResource("ButtonDefault")!
        };
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        bar.Children.Add(ok);
        bar.Children.Add(cancel);

        var list = new ListBox { SelectionMode = SelectionMode.Single };
        int idx = 0;
        foreach (var opt in Options)
        {
            var item = new ListBoxItem { Content = opt.label, Tag = opt.type };
            list.Items.Add(item);
            if (opt.type == ViewType.Table) list.SelectedIndex = idx;
            idx++;
        }
        list.SelectionChanged += (_, _) =>
        {
            if (list.SelectedItem is ListBoxItem li && li.Tag is ViewType t)
                SelectedType = t;
        };

        dock.Children.Add(bar);
        dock.Children.Add(list);
        Content = dock;

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape) { DialogResult = false; Close(); }
        };
    }
}
