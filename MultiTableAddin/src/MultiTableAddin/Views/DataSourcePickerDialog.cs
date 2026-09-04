using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MultiTableAddin.Core;
using ListBox = System.Windows.Controls.ListBox;
using StackPanel = System.Windows.Controls.StackPanel;
using ScrollViewer = System.Windows.Controls.ScrollViewer;
using TextBlock = System.Windows.Controls.TextBlock;
using DockPanel = System.Windows.Controls.DockPanel;

namespace MultiTableAddin.Views;

/// <summary>加载数据前选择具体的工作表 / 超级表（解决多表场景）。</summary>
public class DataSourcePickerDialog : Window
{
    private readonly ListBox _list = new();
    public TableSourceInfo? Selected { get; private set; }

    public DataSourcePickerDialog(System.Collections.Generic.List<TableSourceInfo> sources)
    {
        Title = "选择数据源";
        Width = 440; Height = 380;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = (Brush)FindResource("RegionBrush")!;

        var dock = new DockPanel { Margin = new Thickness(12) };

        var tip = new TextBlock
        {
            Text = "选择要加载的超级表（工作表 / 表）：",
            FontSize = 12,
            Margin = new Thickness(0, 0, 0, 8),
            Foreground = (Brush)FindResource("SecondaryTextBrush")!
        };
        DockPanel.SetDock(tip, Dock.Top);

        var bar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        var ok = new Button { Content = "确定", Width = 88, Style = (Style)FindResource("ButtonPrimary")! };
        ok.Click += (_, _) => { if (Commit()) { DialogResult = true; Close(); } };
        var cancel = new Button { Content = "取消", Width = 88, Margin = new Thickness(10, 0, 0, 0), Style = (Style)FindResource("ButtonDefault")! };
        cancel.Click += (_, _) => { DialogResult = false; Close(); };
        bar.Children.Add(ok); bar.Children.Add(cancel);
        DockPanel.SetDock(bar, Dock.Bottom);

        _list.DisplayMemberPath = "DisplayText";
        _list.ItemsSource = sources;
        var sv = new ScrollViewer { Content = _list, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };

        dock.Children.Add(tip);
        dock.Children.Add(bar);
        dock.Children.Add(sv);
        Content = dock;

        Loaded += (_, _) => { if (_list.HasItems) _list.SelectedIndex = 0; };

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        };
    }

    private bool Commit()
    {
        Selected = _list.SelectedItem as TableSourceInfo;
        return Selected != null;
    }
}
