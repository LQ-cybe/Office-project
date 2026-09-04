using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using MultiTableAddin.Views;

namespace MultiTableAddin.TaskPane;

public partial class TaskPaneWpfView : System.Windows.Controls.UserControl
{
    private bool _suppressAccentPresetSelectionChanged;

    public TaskPaneWpfView()
    {
        InitializeComponent();
        InitializeAccentPresetDemo();
        UpdateTaskPaneStatus();
        StatusTextBlock.Text = "状态：已加载，当前可继续接入业务逻辑。";
    }

    private void OnCopyClick(object sender, RoutedEventArgs e)
    {
        string value = string.IsNullOrWhiteSpace(InputTextBox.Text)
            ? "MultiTableAddin 已加载"
            : InputTextBox.Text.Trim();

        System.Windows.Clipboard.SetText(value);
        StatusTextBlock.Text = "状态：已复制到剪贴板。";
        AddInLog.Write("TaskPane.Copy", value);
    }

    private void OnOpenLogClick(object sender, RoutedEventArgs e)
    {
        string logPath = AddInLog.LogFilePath;

        if (!File.Exists(logPath))
        {
            AddInLog.Write("TaskPane.OpenLog", "LogFileMissing");
            StatusTextBlock.Text = "状态：日志文件还未生成。";
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "/select,\"" + logPath + "\"",
            UseShellExecute = true
        });

        StatusTextBlock.Text = "状态：已打开日志位置。";
    }

    private void OnOpenDialogClick(object sender, RoutedEventArgs e)
    {
        PopupDemoView popupDemoView = new PopupDemoView();
        DemoWindowService.ShowHostedWpfControl(
            "MultiTableAddin - 示例窗口",
            popupDemoView,
            560,
            860);

        StatusTextBlock.Text = "状态：已打开弹出窗体。";
    }

    private void OnRefreshTaskPaneStatusClick(object sender, RoutedEventArgs e)
    {
        UpdateTaskPaneStatus();
        StatusTextBlock.Text = "状态：已刷新当前窗口的窗格状态。";
    }

    private void InitializeAccentPresetDemo()
    {
        _suppressAccentPresetSelectionChanged = true;
        AccentPresetComboBox.Items.Clear();

        foreach (HandyControlRuntime.AccentPresetOption preset in HandyControlRuntime.GetAccentPresetOptions())
        {
            AccentPresetComboBox.Items.Add(new ComboBoxItem
            {
                Content = preset.DisplayName,
                Tag = preset.Key
            });
        }

        AccentPresetComboBox.SelectedIndex = 0;
        AccentPresetHintTextBlock.Text = "这里演示 HandyControl 内置 Accent 预设，可直接让用户试红、橙、金、绿、蓝、紫等内置色。";
        _suppressAccentPresetSelectionChanged = false;
    }

    private void OnAccentPresetSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressAccentPresetSelectionChanged)
        {
            return;
        }

        string presetKey = (AccentPresetComboBox.SelectedItem as ComboBoxItem)?.Tag as string
            ?? HandyControlRuntime.DefaultAccentPresetKey;

        if (HandyControlRuntime.TryApplyAccentPreset(presetKey, out HandyControlRuntime.AccentPresetOption preset))
        {
            AccentPresetHintTextBlock.Text = preset.Description;
            StatusTextBlock.Text = "状态：已切换主题 Accent 为「" + preset.DisplayName + "」。";
            return;
        }

        StatusTextBlock.Text = "状态：切换 Accent 预设失败，请查看日志。";
    }

    private void UpdateTaskPaneStatus()
    {
        TaskPaneStatusSnapshot snapshot = TaskPaneManager.GetStatusSnapshot();
        TaskPaneWindowKeyTextBlock.Text = snapshot.ActiveWindowKey;
        TaskPaneWindowCaptionTextBlock.Text = snapshot.ActiveWindowCaption;
        TaskPaneManagerStateTextBlock.Text = string.Format(
            "已管理窗口数：{0}；{1}",
            snapshot.ManagedWindowCount,
            snapshot.VisibilityText);
    }
}
