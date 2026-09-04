using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ListBox = System.Windows.Controls.ListBox;
using SelectionMode = System.Windows.Controls.SelectionMode;

namespace MultiTableAddin.Views;

/// <summary>简单字段多选对话框，用于看板/画册等视图配置显示字段</summary>
public class FieldPickerDialog : Window
{
    public List<string> SelectedFields { get; } = new();

    public FieldPickerDialog(List<string> allFields, List<string> selected, string title)
    {
        Title = title;
        Width = 360;
        Height = 460;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Loaded += (_, _) => Background = TryFindResource("RegionBrush") as Brush ?? Brushes.White;

        var dock = new DockPanel { Margin = new Thickness(14) };

        var bar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        DockPanel.SetDock(bar, Dock.Bottom);
        var ok = new Button { Content = "确定", Width = 80 };
        ok.Click += (_, _) => { DialogResult = true; Close(); };
        var cancel = new Button { Content = "取消", Width = 80, Margin = new Thickness(10, 0, 0, 0) };
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        Loaded += (_, _) =>
        {
            ok.Style = TryFindResource("ButtonPrimary") as Style;
            cancel.Style = TryFindResource("ButtonDefault") as Style;
        };
        bar.Children.Add(ok);
        bar.Children.Add(cancel);

        var listBox = new ListBox { SelectionMode = SelectionMode.Multiple };

        // 明确选中态样式：选中项用浅蓝底 + 蓝字，避免默认主题下高亮不明显
        var itemStyle = new Style(typeof(ListBoxItem));
        itemStyle.Setters.Add(new Setter(ListBoxItem.PaddingProperty, new Thickness(8, 6, 8, 6)));
        itemStyle.Setters.Add(new Setter(ListBoxItem.BackgroundProperty, Brushes.Transparent));
        itemStyle.Triggers.Add(new Trigger
        {
            Property = ListBoxItem.IsSelectedProperty,
            Value = true,
            Setters =
            {
                new Setter(ListBoxItem.BackgroundProperty, new SolidColorBrush(Color.FromRgb(220, 235, 255))),
                new Setter(ListBoxItem.ForegroundProperty, new SolidColorBrush(Color.FromRgb(20, 90, 200))),
                new Setter(ListBoxItem.FontWeightProperty, FontWeights.SemiBold)
            }
        });
        listBox.ItemContainerStyle = itemStyle;

        // 先创建所有项，保存引用
        var items = new List<ListBoxItem>();
        foreach (var f in allFields)
        {
            var item = new ListBoxItem { Content = f };
            listBox.Items.Add(item);
            items.Add(item);
        }

        // 窗口加载完成后重新应用选中态，确保打开即高亮当前显示字段
        Loaded += (_, _) =>
        {
            foreach (var item in items)
            {
                string? name = item.Content?.ToString();
                if (name != null && selected.Contains(name))
                {
                    item.IsSelected = true;
                    if (!SelectedFields.Contains(name)) SelectedFields.Add(name);
                }
            }
            listBox.Focus();
        };

        listBox.SelectionChanged += (_, _) =>
        {
            SelectedFields.Clear();
            foreach (ListBoxItem item in listBox.SelectedItems)
                SelectedFields.Add(item.Content?.ToString() ?? string.Empty);
        };

        dock.Children.Add(bar);
        dock.Children.Add(listBox);
        Content = dock;

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape) { DialogResult = false; Close(); }
        };
    }
}
