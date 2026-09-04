using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace MultiTableAddin.TaskPane;

public partial class WebTaskPaneWpfView : System.Windows.Controls.UserControl, ITaskPaneCloseAware
{
    private const string DefaultBaiduUrl = "https://www.baidu.com";

    private readonly WebView2 _webView;
    private bool _isInitialized;
    private bool _isClosing;
    private bool _coreEventsAttached;
    private string? _pendingBaiduSearchText;

    public WebTaskPaneWpfView()
    {
        InitializeComponent();

        _webView = new WebView2
        {
            Dock = System.Windows.Forms.DockStyle.Fill
        };

        AddressTextBox.Text = DefaultBaiduUrl;
        DataFolderTextBlock.Text = "数据目录：" + AddInRuntime.GetWebView2DataDirectory();
        AddressTextBox.KeyDown += OnAddressTextBoxKeyDown;
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (_isInitialized || _isClosing)
        {
            return;
        }

        _isInitialized = true;
        await InitializeWebViewAsync();
    }

    private async Task InitializeWebViewAsync()
    {
        if (_isClosing)
        {
            return;
        }

        try
        {
            string userDataFolder = AddInRuntime.GetWebView2DataDirectory();
            Directory.CreateDirectory(userDataFolder);

            WebViewHost.Child = _webView;

            CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: userDataFolder,
                options: null);

            await _webView.EnsureCoreWebView2Async(environment);
            if (_isClosing)
            {
                return;
            }

            AttachCoreEvents();
            // #360 屏蔽 WebView2 浏览器原生右键菜单（主窗口与任务窗格一致）
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = true;

            ErrorBorder.Visibility = Visibility.Collapsed;
            SetStatus("WebView2 初始化完成");
            NavigateToAddress(AddressTextBox.Text);
            AddInLog.Write("WebView2.Initialized", userDataFolder);
        }
        catch (Exception ex)
        {
            if (!_isClosing)
            {
                SetInitializationError(ex);
            }
        }
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        AddInLog.Write("WebView2.Unloaded", "VisualTreeUnloaded");
    }

    private void OnGoClick(object sender, RoutedEventArgs e)
    {
        NavigateToAddress(AddressTextBox.Text);
    }

    private void OnHomeClick(object sender, RoutedEventArgs e)
    {
        AddressTextBox.Text = DefaultBaiduUrl;
        NavigateToAddress(AddressTextBox.Text);
    }

    private void OnRefreshClick(object sender, RoutedEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        try
        {
            if (_webView.CoreWebView2 == null)
            {
                return;
            }

            _webView.Reload();
            SetStatus("正在刷新页面");
            AddInLog.Write("WebView2.Reload", AddressTextBox.Text);
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.Reload.Error", ex.ToString());
        }
    }

    private void OnOpenDataDirectoryClick(object sender, RoutedEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        string path = AddInRuntime.GetWebView2DataDirectory();
        Directory.CreateDirectory(path);

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "\"" + path + "\"",
            UseShellExecute = true
        });

        SetStatus("已打开 WebView2 数据目录");
        AddInLog.Write("WebView2.OpenDataDirectory", path);
    }

    private async void OnSearchA1Click(object sender, RoutedEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        string keyword = ReadActiveCellText("A1");
        CellPreviewTextBlock.Text = string.IsNullOrWhiteSpace(keyword)
            ? "A1：为空"
            : "A1：" + keyword;

        if (string.IsNullOrWhiteSpace(keyword))
        {
            SetStatus("A1 为空，未执行搜索");
            AddInLog.Write("WebView2.SearchA1.Empty", "A1");
            return;
        }

        if (_webView.CoreWebView2 == null)
        {
            _pendingBaiduSearchText = keyword;
            SetStatus("WebView2 尚未初始化，已暂存 A1 搜索");
            AddInLog.Write("WebView2.SearchA1.Pending", keyword);
            return;
        }

        try
        {
            _pendingBaiduSearchText = keyword;
            if (!IsBaiduAddress(_webView.Source?.ToString()))
            {
                AddressTextBox.Text = DefaultBaiduUrl;
                NavigateToAddress(DefaultBaiduUrl);
                return;
            }

            await ExecutePendingBaiduSearchAsync();
        }
        catch (Exception ex)
        {
            SetStatus("A1 搜索执行失败");
            AddInLog.Write("WebView2.SearchA1.Error", ex.ToString());
        }
    }

    private void OnAddressTextBoxKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            NavigateToAddress(AddressTextBox.Text);
        }
    }

    private void CoreWebView2_NavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        SetStatus("正在打开 " + e.Uri);
    }

    private void CoreWebView2_NewWindowRequested(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
    {
        if (_isClosing)
        {
            e.Handled = true;
            return;
        }

        try
        {
            string target = string.IsNullOrWhiteSpace(e.Uri) ? DefaultBaiduUrl : e.Uri;
            e.Handled = true;
            AddressTextBox.Text = target;
            _webView.CoreWebView2.Navigate(target);
            SetStatus("已拦截新窗口并在当前页打开");
            AddInLog.Write("WebView2.NewWindowRedirected", target);
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.NewWindowRedirect.Error", ex.ToString());
        }
    }

    private void CoreWebView2_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (_isClosing)
        {
            return;
        }

        string currentAddress = _webView.Source?.ToString() ?? AddressTextBox.Text;
        AddressTextBox.Text = currentAddress;
        SetStatus(e.IsSuccess
            ? "页面加载完成"
            : "页面加载失败，请查看 logs 日志");

        if (e.IsSuccess && !string.IsNullOrWhiteSpace(_pendingBaiduSearchText) && IsBaiduAddress(currentAddress))
        {
            _ = ExecutePendingBaiduSearchAsync();
        }
    }

    private void NavigateToAddress(string? rawAddress)
    {
        if (_isClosing)
        {
            return;
        }

        if (_webView.CoreWebView2 == null)
        {
            SetStatus("WebView2 尚未完成初始化");
            return;
        }

        string target = string.IsNullOrWhiteSpace(rawAddress)
            ? DefaultBaiduUrl
            : rawAddress.Trim();

        if (!Uri.TryCreate(target, UriKind.Absolute, out _))
        {
            target = "https://" + target;
        }

        try
        {
            _webView.CoreWebView2.Navigate(target);
            SetStatus("已发起导航");
            AddInLog.Write("WebView2.Navigate", target);
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.Navigate.Error", ex.ToString());
        }
    }

    private async Task ExecutePendingBaiduSearchAsync()
    {
        if (_isClosing || _webView.CoreWebView2 == null || string.IsNullOrWhiteSpace(_pendingBaiduSearchText))
        {
            return;
        }

        string keyword = _pendingBaiduSearchText;
        _pendingBaiduSearchText = null;
        string jsKeyword = JsonSerializer.Serialize(keyword);
        string script = @"
(() => {
  const keyword = " + jsKeyword + @";
  const input = document.querySelector('#kw, input[name=""wd""]');
  if (!input) return 'input-not-found';
  input.focus();
  input.value = keyword;
  input.dispatchEvent(new Event('input', { bubbles: true }));
  input.dispatchEvent(new Event('change', { bubbles: true }));
  const button = document.querySelector('#su, input[type=""submit""][value=""百度一下""], button[type=""submit""]');
  if (button) {
    button.click();
    return 'clicked';
  }
  const form = input.form || document.querySelector('form');
  if (form) {
    form.submit();
    return 'submitted';
  }
  return 'button-not-found';
})();";

        try
        {
            string result = await _webView.CoreWebView2.ExecuteScriptAsync(script);
            SetStatus("已把 A1 内容写入百度并触发搜索");
            AddInLog.Write("WebView2.SearchA1.Executed", keyword + " | Result=" + result);
        }
        catch (Exception ex)
        {
            SetStatus("百度自动搜索失败，请查看日志");
            AddInLog.Write("WebView2.SearchA1.Execute.Error", ex.ToString());
        }
    }

    void ITaskPaneCloseAware.PrepareForClose(TaskPaneCloseContext context)
    {
        if (_isClosing)
        {
            return;
        }

        _isClosing = true;
        _pendingBaiduSearchText = null;
        Loaded -= OnLoaded;
        Unloaded -= OnUnloaded;
        AddressTextBox.KeyDown -= OnAddressTextBoxKeyDown;
        DetachCoreEvents();

        try
        {
            WebViewHost.Child = null;
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.Close.DetachHost.Error", ex.ToString());
        }

        if (context.AvoidAggressiveDispose)
        {
            AddInLog.Write("WebView2.Close.Deferred", "Source=" + context.Source + "; Reason=" + context.Reason);
            return;
        }

        try
        {
            _webView.Stop();
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.Close.Stop.Error", ex.ToString());
        }

        try
        {
            _webView.Dispose();
            AddInLog.Write("WebView2.Close.Disposed", "Source=" + context.Source + "; Reason=" + context.Reason);
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.Close.Dispose.Error", ex.ToString());
        }
    }

    private void AttachCoreEvents()
    {
        if (_coreEventsAttached || _webView.CoreWebView2 == null)
        {
            return;
        }

        _webView.CoreWebView2.NewWindowRequested += CoreWebView2_NewWindowRequested;
        _webView.CoreWebView2.NavigationStarting += CoreWebView2_NavigationStarting;
        _webView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
        _coreEventsAttached = true;
    }

    private void DetachCoreEvents()
    {
        if (!_coreEventsAttached || _webView.CoreWebView2 == null)
        {
            return;
        }

        try
        {
            _webView.CoreWebView2.NewWindowRequested -= CoreWebView2_NewWindowRequested;
            _webView.CoreWebView2.NavigationStarting -= CoreWebView2_NavigationStarting;
            _webView.CoreWebView2.NavigationCompleted -= CoreWebView2_NavigationCompleted;
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.Close.DetachEvents.Error", ex.ToString());
        }

        _coreEventsAttached = false;
    }

    private void SetStatus(string text)
    {
        StatusTextBlock.Text = "状态：" + text;
        AddressTextBox.ToolTip = text;
    }

    private void SetInitializationError(Exception ex)
    {
        AddInLog.Write("WebView2.Initialize.Error", ex.ToString());
        SetStatus("WebView2 初始化失败");
        ErrorTextBlock.Text = "WebView2 初始化失败，请确认机器已安装 WebView2 Runtime，且 dist 目录已带上 runtimes\\win-x64 或 win-x86 下的 WebView2Loader.dll。"
            + Environment.NewLine
            + Environment.NewLine
            + ex;
        ErrorBorder.Visibility = Visibility.Visible;
    }

    private static bool IsBaiduAddress(string? address)
    {
        if (!Uri.TryCreate(address, UriKind.Absolute, out Uri? uri))
        {
            return false;
        }

        return uri.Host.IndexOf("baidu.com", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string ReadActiveCellText(string address)
    {
        try
        {
            dynamic? application = ExcelDnaUtil.Application;
            dynamic? activeSheet = application?.ActiveSheet;
            dynamic? range = activeSheet?.Range[address];
            object? value = range?.Value2;
            return Convert.ToString(value)?.Trim() ?? string.Empty;
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.ReadCell.Error", address + " | " + ex);
            return string.Empty;
        }
    }
}
