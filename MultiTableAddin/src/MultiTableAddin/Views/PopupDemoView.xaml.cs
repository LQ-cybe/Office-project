using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using ExcelDna.Integration;

namespace MultiTableAddin.Views;

public partial class PopupDemoView : System.Windows.Controls.UserControl, IRequestClose
{
    private bool _suspendInteractiveToggle;
    private bool _interactiveWasForcedOff;

    public PopupDemoView()
    {
        _suspendInteractiveToggle = true;
        InitializeComponent();
        _suspendInteractiveToggle = false;
        Loaded += PopupDemoView_Loaded;
        Unloaded += PopupDemoView_Unloaded;
        StatusTextBlock.Text = "状态：窗口已打开。若高版本运行时下输入不稳，可先尝试 Interactive 开关。";
    }

    public event EventHandler? RequestClose;

    internal void FocusPrimaryInput()
    {
        UserNameTextBox.Focus();
        Keyboard.Focus(UserNameTextBox);
    }

    private void PopupDemoView_Loaded(object sender, RoutedEventArgs e)
    {
        RefreshInteractiveToggleState();
    }

    private void PopupDemoView_Unloaded(object sender, RoutedEventArgs e)
    {
        RestoreInteractiveIfNeeded();
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        string userName = string.IsNullOrWhiteSpace(UserNameTextBox.Text) ? "未命名用户" : UserNameTextBox.Text.Trim();
        string license = string.IsNullOrWhiteSpace(LicenseTextBox.Text) ? "未填写" : LicenseTextBox.Text.Trim();

        string message = $"User={userName}; License={license}";
        AddInLog.Write("Dialog.Save", message);
        StatusTextBlock.Text = "状态：已写入日志。";
    }

    private void OnOpenLogClick(object sender, RoutedEventArgs e)
    {
        string logPath = AddInLog.LogFilePath;
        if (!File.Exists(logPath))
        {
            StatusTextBlock.Text = "状态：日志文件尚未生成。";
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

    private void OnInteractiveToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suspendInteractiveToggle)
        {
            return;
        }

        bool shouldEnable = InteractiveToggleButton.IsChecked == true;
        if (TrySetApplicationInteractive(shouldEnable))
        {
            _interactiveWasForcedOff = !shouldEnable;
            StatusTextBlock.Text = shouldEnable
                ? "状态：已恢复 Excel 交互，可返回工作簿浏览和复制内容。"
                : "状态：已临时关闭 Excel 交互，用于缓解输入被单元格抢走；若仍不稳，请考虑改用 ShowDialog。";
            AddInLog.Write("Dialog.Interactive.Toggle", "SetTo=" + shouldEnable);
            return;
        }

        StatusTextBlock.Text = "状态：切换 Application.Interactive 失败，请查看日志。";
        RefreshInteractiveToggleState();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e)
    {
        RestoreInteractiveIfNeeded();
        RequestClose?.Invoke(this, EventArgs.Empty);
    }

    private void RefreshInteractiveToggleState()
    {
        bool? currentValue = TryGetApplicationInteractive();
        if (currentValue == null)
        {
            StatusTextBlock.Text = "状态：无法读取 Excel 交互状态，请查看日志。";
            return;
        }

        _suspendInteractiveToggle = true;
        InteractiveToggleButton.IsChecked = currentValue.Value;
        _suspendInteractiveToggle = false;

        StatusTextBlock.Text = currentValue.Value
            ? "状态：Excel 当前可交互。若文本输入掉到单元格，可临时关闭 Interactive。"
            : "状态：Excel 当前已被临时锁定交互。需要回表查看时请重新打开 Interactive。";
    }

    private static bool? TryGetApplicationInteractive()
    {
        try
        {
            dynamic? application = ExcelDnaUtil.Application;
            if (application == null)
            {
                return null;
            }

            return Convert.ToBoolean(application.Interactive);
        }
        catch (Exception ex)
        {
            AddInLog.Write("Dialog.Interactive.Read.Error", ex.ToString());
            return null;
        }
    }

    private static bool TrySetApplicationInteractive(bool value)
    {
        try
        {
            dynamic? application = ExcelDnaUtil.Application;
            if (application == null)
            {
                return false;
            }

            application.Interactive = value;
            return true;
        }
        catch (Exception ex)
        {
            AddInLog.Write("Dialog.Interactive.Set.Error", "SetTo=" + value + " | " + ex);
            return false;
        }
    }

    private void RestoreInteractiveIfNeeded()
    {
        if (!_interactiveWasForcedOff)
        {
            return;
        }

        if (TrySetApplicationInteractive(true))
        {
            _interactiveWasForcedOff = false;
            AddInLog.Write("Dialog.Interactive.Restore", "SetTo=True");
        }
    }
}
