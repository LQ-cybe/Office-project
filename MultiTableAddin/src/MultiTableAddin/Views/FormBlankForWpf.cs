using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using ExcelDna.Integration;

namespace MultiTableAddin.Views;

public partial class FormBlankForWpf : Form
{
    private ElementHost? _elementHost;
    private ModelessKeyboardMessageFilter? _keyboardMessageFilter;

    public FormBlankForWpf(int height, int width)
    {
        InitializeComponent();

        BackColor = System.Drawing.Color.White;

        float scalingFactor = GetScalingFactor();
        Width = (int)(width * scalingFactor);
        Height = (int)((height + 30) * scalingFactor);
    }

    public Action? ActionFormClosing { get; set; }

    public System.Windows.Controls.UserControl? WpfUserControl { get; set; }

    private void FormBlankForWpf_Load(object? sender, EventArgs e)
    {
        try
        {
            if (WpfUserControl == null)
            {
                return;
            }

            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = WpfUserControl
            };

            Controls.Add(_elementHost);

            if (WpfUserControl is IRequestClose requestCloseControl)
            {
                requestCloseControl.RequestClose += WpfUserControl_RequestClose;
            }

            EnsureKeyboardInteropFilter();
            BeginInvoke(new Action(ActivateHostedContent));
        }
        catch (Exception ex)
        {
            AddInLog.Write("Dialog.Host.Error", ex.ToString());
            MessageBox.Show("WPF 窗体加载失败，请查看日志：" + AddInLog.LogFilePath, "MultiTableAddin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void WpfUserControl_RequestClose(object? sender, EventArgs e)
    {
        Close();
    }

    public float GetScalingFactor()
    {
        using Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
        return graphics.DpiX / 96f;
    }

    private void FormBlankForWpf_FormClosing(object? sender, FormClosingEventArgs e)
    {
        try
        {
            if (WpfUserControl is IRequestClose requestCloseControl)
            {
                requestCloseControl.RequestClose -= WpfUserControl_RequestClose;
            }

            if (_keyboardMessageFilter != null)
            {
                Application.RemoveMessageFilter(_keyboardMessageFilter);
                _keyboardMessageFilter = null;
            }

            ActionFormClosing?.Invoke();
        }
        catch
        {
        }
    }

    private void FormBlankForWpf_Shown(object? sender, EventArgs e)
    {
        ActivateHostedContent();
    }

    internal void ActivateHostedContent()
    {
        try
        {
            Activate();
            BringToFront();
            Focus();
            _elementHost?.Focus();

            if (WpfUserControl == null)
            {
                return;
            }

            WpfUserControl.Dispatcher.BeginInvoke(new Action(() =>
            {
                WpfUserControl.Focus();
                Keyboard.Focus(WpfUserControl);
                WpfUserControl.MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
            }));
        }
        catch (Exception ex)
        {
            AddInLog.Write("Dialog.Focus.Error", ex.ToString());
        }
    }

    private void EnsureKeyboardInteropFilter()
    {
        if (_keyboardMessageFilter != null)
        {
            return;
        }

        _keyboardMessageFilter = new ModelessKeyboardMessageFilter(this);
        Application.AddMessageFilter(_keyboardMessageFilter);
    }
}

internal static class DemoWindowService
{
    internal static DialogResult ShowHostedWpfDialog(string title, System.Windows.Controls.UserControl userControl, int height, int width)
    {
        FormBlankForWpf form = CreateHostedForm(title, userControl, height, width);
        IWin32Window? owner = TryCreateExcelOwnerWindow();

        // 模态窗口是高版本运行时下非模态焦点不稳时的兜底方案。
        AddInLog.Write("Dialog.ShowHostedForm.Modal", title);
        return owner != null
            ? form.ShowDialog(owner)
            : form.ShowDialog();
    }

    internal static FormBlankForWpf ShowHostedWpfControl(string title, System.Windows.Controls.UserControl userControl, int height, int width)
    {
        FormBlankForWpf form = CreateHostedForm(title, userControl, height, width);
        IWin32Window? owner = TryCreateExcelOwnerWindow();

        // 当前模板默认优先保留非模态体验，但目标运行时升高后仍需做焦点实机验证。
        if (owner != null)
        {
            form.Show(owner);
        }
        else
        {
            form.Show();
        }
        form.Location = GetFormCenterLocation(form);
        form.ActivateHostedContent();
        return form;
    }

    private static FormBlankForWpf CreateHostedForm(string title, System.Windows.Controls.UserControl userControl, int height, int width)
    {
        return new FormBlankForWpf(height, width)
        {
            Text = title,
            WpfUserControl = userControl,
            ShowInTaskbar = false
        };
    }

    private static System.Drawing.Point GetFormCenterLocation(Form form)
    {
        System.Drawing.Rectangle workingArea = Screen.FromPoint(Control.MousePosition).WorkingArea;
        int x = workingArea.Left + Math.Max(0, (workingArea.Width - form.Width) / 2);
        int y = workingArea.Top + Math.Max(0, (workingArea.Height - form.Height) / 2);
        return new System.Drawing.Point(x, y);
    }

    private static IWin32Window? TryCreateExcelOwnerWindow()
    {
        try
        {
            IntPtr handle = ExcelDnaUtil.WindowHandle;
            if (handle == IntPtr.Zero)
            {
                return null;
            }

            return new Win32Window(handle);
        }
        catch (Exception ex)
        {
            AddInLog.Write("Dialog.OwnerWindow.Error", ex.ToString());
            return null;
        }
    }
}

internal sealed class Win32Window : IWin32Window
{
    internal Win32Window(IntPtr handle)
    {
        Handle = handle;
    }

    public IntPtr Handle { get; }
}

internal sealed class ModelessKeyboardMessageFilter : IMessageFilter
{
    private const int WmKeyDown = 0x0100;
    private const int WmKeyUp = 0x0101;
    private const int WmChar = 0x0102;
    private const int WmSysKeyDown = 0x0104;
    private const int WmSysKeyUp = 0x0105;
    private const int WmSysChar = 0x0106;

    private readonly FormBlankForWpf _form;

    internal ModelessKeyboardMessageFilter(FormBlankForWpf form)
    {
        _form = form;
    }

    public bool PreFilterMessage(ref Message m)
    {
        if (!_form.Visible || _form.IsDisposed)
        {
            return false;
        }

        if (!IsKeyboardMessage(m.Msg))
        {
            return false;
        }

        bool formIsActive = Form.ActiveForm == _form || _form.ContainsFocus;
        if (!formIsActive)
        {
            return false;
        }

        return _form.PreProcessMessage(ref m);
    }

    private static bool IsKeyboardMessage(int messageId)
    {
        return messageId == WmKeyDown
            || messageId == WmKeyUp
            || messageId == WmChar
            || messageId == WmSysKeyDown
            || messageId == WmSysKeyUp
            || messageId == WmSysChar;
    }
}
