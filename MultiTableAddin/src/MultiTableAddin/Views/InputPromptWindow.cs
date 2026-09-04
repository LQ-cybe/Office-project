using System.Windows;
using TextBox = System.Windows.Controls.TextBox;
using TextBlock = System.Windows.Controls.TextBlock;
using StackPanel = System.Windows.Controls.StackPanel;

namespace MultiTableAddin.Views;

/// <summary>通用文本输入弹窗（用于视图重命名等场景）</summary>
public class InputPromptWindow : Window
{
    private readonly TextBox _textBox;

    /// <summary>用户输入的结果；取消时为 null</summary>
    public string? ResultText { get; private set; }

    public InputPromptWindow(string title, string prompt, string defaultValue = "")
    {
        Title = title;
        Width = 360;
        Height = 150;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;

        var panel = new StackPanel { Margin = new Thickness(16) };

        panel.Children.Add(new TextBlock
        {
            Text = prompt,
            Margin = new Thickness(0, 0, 0, 8),
            TextWrapping = TextWrapping.Wrap
        });

        _textBox = new TextBox { Text = defaultValue, Padding = new Thickness(6) };
        panel.Children.Add(_textBox);

        var bar = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };

        var ok = new Button { Content = "确定", Width = 72, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
        var cancel = new Button { Content = "取消", Width = 72, IsCancel = true };
        ok.Click += (_, _) => { ResultText = _textBox.Text.Trim(); DialogResult = true; };
        cancel.Click += (_, _) => { DialogResult = false; };
        bar.Children.Add(ok);
        bar.Children.Add(cancel);

        panel.Children.Add(bar);
        Content = panel;

        Loaded += (_, _) => { _textBox.Focus(); _textBox.SelectAll(); };
    }
}
