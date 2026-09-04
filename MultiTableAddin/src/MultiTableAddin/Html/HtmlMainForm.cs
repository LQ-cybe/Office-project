using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using MultiTableAddin.Core;

namespace MultiTableAddin.Html;

/// <summary>
/// 单 WebView2 主窗口：托管内嵌的 HTML/CSS/JS 单页应用（SPA）。
/// 所有多维表界面（侧栏 + 视图区 + 弹窗）均由 HTML 绘制；
/// C# 仅通过 WebMessageReceived 桥接提供 Excel 读写与配置存取能力。
/// HTML 作为内嵌资源随 DLL 一起发布，UI 版本与 DLL 锁死，避免“改了不生效”。
/// </summary>
public class HtmlMainForm : Form
{
    private static HtmlMainForm? _instance;
    public static HtmlMainForm? Instance => _instance;

    private readonly WebView2 _webView;
    private bool _coreReady;
    private readonly SynchronizationContext _uiContext;

    // #402 甘特手动保存模式：JS 推送的“未保存可视化调整”脏标记，供关闭软件前拦截确认
    private bool _visualEditDirty;
    private int _visualEditCount;
    private bool _forceClosing;

    // 当前打开的表上下文
    private string _workbookPath = string.Empty;
    private string _sheetName = string.Empty;
    private string _tableName = string.Empty;
    private DataTableModel? _table;
    private ViewConfigFile? _config;

    // 跟随 Excel 主窗口最小化/还原而自动隐藏/显示（替代“拥有窗口”方案，避免阻塞 Excel 任务栏）
    private IntPtr _excelHwnd;
    private IntPtr _winEventHook;
    private IntPtr _fgHook;
    private uint _excelPid;
    private bool _topMostWanted;
    private WinEventDelegate? _winEventDelegate;
    private const uint EVENT_SYSTEM_MINIMIZESTART = 0x0016;
    private const uint EVENT_SYSTEM_MINIMIZEEND = 0x0017;
    private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
    private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
    private const uint WINEVENT_SKIPOWNPROCESS = 0x0002;
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventTime);
    [DllImport("user32.dll")]
    private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);
    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters = { new JsonStringEnumConverter() }
    };

    public HtmlMainForm()
    {
        _instance = this;
        _uiContext = SynchronizationContext.Current ?? new SynchronizationContext();
        Text = "多维表分析";
        Width = 1200;
        Height = 780;
        WindowState = FormWindowState.Maximized;
        BackColor = System.Drawing.Color.White;
        // 不出现在任务栏，避免与 Excel 各自占一个任务栏按钮；随 Excel 最小化一起隐藏的行为由 WinEvent 钩子实现
        ShowInTaskbar = false;
        try { Icon = LoadWindowIcon(); } catch { }
        _webView = new WebView2 { Dock = DockStyle.Fill };
        Controls.Add(_webView);
        Load += async (s, e) => await InitializeAsync();
        // 跟随 Excel 主窗口最小化/还原，自动隐藏/显示本窗口（替代“拥有窗口”以避免阻塞任务栏）
        Load += (s, e) => AttachExcelMinimizeFollow();
        Load += (s, e) => AttachExcelForegroundWatch();
        // 打开时短暂置顶，确保窗口显示在最上方；随后恢复“不强制置顶”（默认行为）。
        // HTML 右上角的“置顶”按钮可在运行时通过 setTopMost 切换强制置顶。
        Load += (s, e) =>
        {
            try { TopMost = true; BringToFront(); Activate(); TopMost = false; }
            catch { }
        };
        FormClosing += (s, e) =>
        {
            // #402 关闭软件前，若有未保存的甘特手动调整，先弹确认；取消则中止关闭
            if (!_forceClosing && _visualEditDirty)
            {
                e.Cancel = true;
                _ = HandlePendingVisualEditsOnClose();
                return;
            }
            _instance = null;
            try { if (_winEventHook != IntPtr.Zero) UnhookWinEvent(_winEventHook); } catch { }
            try { if (_fgHook != IntPtr.Zero) UnhookWinEvent(_fgHook); } catch { }
            try { _webView?.Dispose(); } catch { }
        };
    }

    /// <summary>
    /// 监听 Excel 主窗口的最小化/还原事件：
    /// - 当 Excel 最小化时，本窗口自动隐藏（保持“随 Excel 一起隐藏”的体验）；
    /// - 当 Excel 还原时，本窗口自动重新显示。
    /// 由于本窗口不再以 Excel 为“拥有窗口”（form.Show 不带 owner），不会阻塞 Excel 任务栏的最小化/切换。
    /// </summary>
    private void AttachExcelMinimizeFollow()
    {
        try
        {
            _excelHwnd = ExcelDnaUtil.WindowHandle;
            if (_excelHwnd == IntPtr.Zero) return;
            uint pid;
            GetWindowThreadProcessId(_excelHwnd, out pid);
            _excelPid = pid;
            _winEventDelegate = OnWinEvent;
            _winEventHook = SetWinEventHook(
                EVENT_SYSTEM_MINIMIZESTART, EVENT_SYSTEM_MINIMIZEEND,
                IntPtr.Zero, _winEventDelegate, pid, 0,
                WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
            AddInLog.Write("HtmlMainForm.MinimizeFollow.Attach",
                _winEventHook != IntPtr.Zero ? "ok hwnd=" + _excelHwnd : "hook failed");
        }
        catch { }
    }

    private void OnWinEvent(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventTime)
    {
        // 前台窗口切换：置顶仅对 Excel（或本窗口）生效，切到其它软件时自动取消置顶，不盖住其它程序
        if (eventType == EVENT_SYSTEM_FOREGROUND)
        {
            if (_topMostWanted)
            {
                uint fpid; GetWindowThreadProcessId(hwnd, out fpid);
                bool excelForeground = (fpid == _excelPid);
                PostToUi(() => { try { TopMost = excelForeground; } catch { } });
            }
            return;
        }
        if (hwnd != _excelHwnd || idObject != 0) return; // 仅关心 Excel 主窗口本身的最小化/还原
        if (eventType == EVENT_SYSTEM_MINIMIZESTART)
        {
            PostToUi(() => { if (!IsDisposed && Visible) Visible = false; });
        }
        else if (eventType == EVENT_SYSTEM_MINIMIZEEND)
        {
            PostToUi(() => { if (!IsDisposed) Visible = true; });
        }
    }

    /// <summary>
    /// 监听“前台窗口”切换（系统级）：当置顶开启且用户切到其它软件时，自动取消本窗口置顶，
    /// 使其只浮在 Excel 之上（置顶“仅限于 Excel 窗口有效”），不盖住浏览器等其它程序。
    /// </summary>
    private void AttachExcelForegroundWatch()
    {
        try
        {
            if (_excelHwnd == IntPtr.Zero) _excelHwnd = ExcelDnaUtil.WindowHandle;
            if (_excelHwnd == IntPtr.Zero) return;
            if (_excelPid == 0) { uint p; GetWindowThreadProcessId(_excelHwnd, out p); _excelPid = p; }
            _winEventDelegate = OnWinEvent;
            // 系统级前台钩子（idProcess=0），由 OnWinEvent 内部按进程判断
            _fgHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
                IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
        }
        catch { }
    }

    private async Task InitializeAsync()
    {
        try
        {
            string userData = Path.Combine(Path.GetTempPath(), "MultiTableAddin", "webview2");
            Directory.CreateDirectory(userData);
            var env = await CoreWebView2Environment.CreateAsync(null, userData, null);
            await _webView.EnsureCoreWebView2Async(env);
            _webView.CoreWebView2.WebMessageReceived += OnWebMessage;
            // #360 屏蔽 WebView2 浏览器原生右键菜单（含“返回/刷新/检查”等），软件自身的自定义右键菜单仍正常
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
            _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            // 显式开启 Web 消息通道（window.chrome.webview 依赖此开关），避免某些 WebView2 版本下桥接对象不可用
            _webView.CoreWebView2.Settings.IsWebMessageEnabled = true;
            _webView.CoreWebView2.NavigateToString(LoadEmbeddedHtml());
            _coreReady = true;
            AddInLog.Write("HtmlMainForm.Init.Ok", "build=" + BuildInfo.Time + " version=" + BuildInfo.Version + " dll=" + GetDllPath());
        }
        catch (Exception ex)
        {
            AddInLog.Write("HtmlMainForm.Init.Error", ex.ToString());
            try
            {
                // 初始化失败时，把错误直接渲染进 WebView2，而不是只弹一个可能被忽略的 MessageBox
                if (_webView.CoreWebView2 != null)
                {
                    string safe = ex.Message.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
                    _webView.CoreWebView2.NavigateToString(
                        "<!doctype html><html><body style='font-family:sans-serif;padding:24px;color:#333'>" +
                        "<h3>WebView2 初始化失败</h3><pre style='white-space:pre-wrap'>" + safe + "</pre>" +
                        "<p>请确认：① 已安装 <b>WebView2 Runtime</b>；② dist 目录随插件部署（含 runtimes\\win-x64 或 win-x86 下的 WebView2Loader.dll）。</p>" +
                        "</body></html>");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("WebView2 初始化失败：" + ex.Message, "多维表",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
            catch { }
        }
    }

    // 构建信息：DLL 的“最后写入时间”即编译/部署时间，随 DLL 一起变化，
    // 注入到 HTML 后即使数据桥接失败也能在加载页/关于窗口显示，便于核对版本。
    private static class BuildInfo
    {
        static BuildInfo()
        {
            try
            {
                string dll = GetDllPath();
                Time = File.Exists(dll) ? File.GetLastWriteTime(dll).ToString("yyyy-MM-dd HH:mm:ss") : "未知";
                Version = AppVersion.DisplayText ?? "未知";
            }
            catch
            {
                Time = "未知"; Version = "未知";
            }
        }
        public static string Time { get; } = "未知";
        public static string Version { get; } = "未知";
    }

    private static System.Drawing.Icon LoadWindowIcon()
    {
        var asm = Assembly.GetExecutingAssembly();
        using Stream? stream = asm.GetManifestResourceStream("MultiTableAddin.Resources.AppIcon.ico");
        if (stream == null)
        {
            AddInLog.Write("HtmlMainForm.Icon.Missing", "MultiTableAddin.Resources.AppIcon.ico");
            throw new InvalidOperationException("未找到窗口图标资源");
        }
        return new System.Drawing.Icon(stream);
    }

    private static string GetDllPath()
    {
        try
        {
            var asm = Assembly.GetExecutingAssembly();
            string name = asm.GetName().Name ?? "MultiTableAddin";

            // 常规加载方式：Assembly.Location 有效
            string loc = asm.Location;
            if (!string.IsNullOrEmpty(loc) && File.Exists(loc)) return loc;

            // Excel-DNA 通常从内存加载插件 DLL，Location 为空；
            // 但 XLL 一定在磁盘上，且 XLL 与 DLL 位于同一目录（dist/files）。
            string addInDir = AddInRuntime.GetAddInDirectory();
            string dll = Path.Combine(addInDir, name + ".dll");
            if (File.Exists(dll)) return dll;
        }
        catch { }
        return "";
    }

    private static string LoadEmbeddedHtml()
    {
        var asm = Assembly.GetExecutingAssembly();
        using var stream = asm.GetManifestResourceStream("MultiTableAddin.Html.index");
        if (stream == null) return "<!doctype html><html><body>未找到 HTML 资源（MultiTableAddin.Html.index）</body></html>";
        using var reader = new StreamReader(stream);
        string html = reader.ReadToEnd();

        // 注入构建信息脚本（window.__mtBuild），供加载页与关于窗口直接读取。
        string dll = GetDllPath().Replace("\\", "\\\\");
        string script = "<script>window.__mtBuild={time:\"" + BuildInfo.Time + "\",version:\"" + BuildInfo.Version + "\",dll:\"" + dll + "\"};</script>";
        html = html.Replace("<!--MT_BUILD-->", script);
        return html;
    }

    private void OnWebMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        string json = e.TryGetWebMessageAsString();
        if (string.IsNullOrWhiteSpace(json)) return;
        AddInLog.Write("HtmlMainForm.Message.Received", json.Length + " chars");

        string? id = null;
        string method = "";
        string paramsJson = "{}";
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            id = root.TryGetProperty("id", out var idEl) ? idEl.GetString() : null;
            method = root.TryGetProperty("method", out var mEl) ? mEl.GetString() ?? "" : "";
            if (root.TryGetProperty("params", out var pEl))
                paramsJson = pEl.GetRawText();
        }
        catch (Exception ex)
        {
            AddInLog.Write("HtmlMainForm.Message.ParseError", ex.ToString());
            return;
        }

        AddInLog.Write("HtmlMainForm.Message.Received", "method=" + method + " id=" + (id ?? "null"));

        // 关键修复：WebView2 的 WebMessageReceived 虽在主线程触发，但属于“非宏上下文”。
        // 直接（同步）访问 Excel 对象模型会触发 COM 重入死锁——主线程卡在等待 COM 返回，
        // 而 Excel 又因调用栈未清空而拒绝放行，导致 C# 永远不回消息、JS 一直 pending（页面卡在“正在加载”）。
        // 用 ExcelAsyncUtil.QueueAsMacro 把 Excel 访问排到宏上下文执行，立即返回、不再阻塞主线程。
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // Excel 访问必须在宏上下文完成；但回传 WebView2 消息必须切回 UI 上下文，
                // 否则在宏执行期间直接 PostWebMessageAsJson 会被消息泵吞掉，JS 端永远收不到而超时。
                try
                {
                    using var pdoc = JsonDocument.Parse(paramsJson);
                    object result = Dispatch(method, pdoc.RootElement);
                    PostToUi(() =>
                    {
                        Send(id, true, result, null);
                        AddInLog.Write("HtmlMainForm.Message.Dispatched", method);
                    });
                }
                catch (Exception ex)
                {
                    AddInLog.Write("HtmlMainForm.Message.DispatchError", method + " :: " + ex);
                    PostToUi(() => Send(id, false, null, ex.Message));
                }
            });
            AddInLog.Write("HtmlMainForm.Message.Queued", method);
        }
        catch (Exception ex)
        {
            // QueueAsMacro 同步失败（极少）：切回 UI 上下文回报错误，避免 JS 永久 pending 卡在“正在加载”
            AddInLog.Write("HtmlMainForm.Message.QueueError", method + " :: " + ex);
            PostToUi(() => Send(id, false, null, "宏排队失败：" + ex.Message));
        }
    }

    /// <summary>
    /// 把操作切回创建 HtmlMainForm 的 WinForms/UI 线程执行。
    /// 在 Excel 宏上下文中直接调用 Send 会导致 WebView2 消息被吞；
    /// 通过 BeginInvoke/同步上下文 post 到消息泵，等宏结束后再回传，JS 才能收到。
    /// </summary>
    private void PostToUi(Action action)
    {
        try
        {
            if (IsHandleCreated)
            {
                this.BeginInvoke(action);
                return;
            }
        }
        catch { }
        _uiContext.Post(_ => action(), null);
    }

    private void Send(string? id, bool ok, object? data, string? error)
    {
        try
        {
            string outJson = JsonSerializer.Serialize(new { id, ok, data, error }, JsonOpts);
            _webView?.CoreWebView2?.PostWebMessageAsJson(outJson);
            AddInLog.Write("HtmlMainForm.Message.Sent", $"id={id ?? "null"} ok={ok} len={outJson.Length}");
        }
        catch (Exception ex)
        {
            AddInLog.Write("HtmlMainForm.Send.Error", ex.ToString());
        }
    }

    /// <summary>
    /// #363/#368/#375/#376 仪表盘导出：把当前 WebView2 截图按「图表内容区」裁剪后保存为 PNG 或复制到剪贴板。
    /// #368/#375 裁剪区域不含左侧视图切换按钮与顶部功能按钮；#375 左上各内缩 15px、宽高按滚动容器完整内容捕获。
    /// #376 已移除 PDF 导出（手写 PDF 为栅格位图、非矢量，依需求删除）。
    /// #379 改为 CDP Page.captureScreenshot(captureBeyondViewport:true) 截整页（绕过 CapturePreviewAsync 仅截视口的限制，修复滚动区外图表被截断）；导出前 JS 解除 .dash-wrap 滚动锁定、导出后还原；保留 CapturePreviewAsync 作兜底。
    /// 截图在 UI 线程异步执行（CapturePreviewAsync 需要 WebView2 就绪），
    /// 不阻塞 Excel 宏上下文；结果通过 NotifyToast 回传给 JS 提示。
    /// </summary>
    private void ExportDashboard(JsonElement pars)
    {
        bool clipboard = false;
        int x = 0, y = 0, w = 0, h = 0;
        if (pars.TryGetProperty("clipboard", out var cEl)) clipboard = cEl.ValueKind == JsonValueKind.True;
        if (pars.TryGetProperty("x", out var xEl)) { if (!xEl.TryGetInt32(out x)) x = 0; }
        if (pars.TryGetProperty("y", out var yEl)) { if (!yEl.TryGetInt32(out y)) y = 0; }
        if (pars.TryGetProperty("w", out var wEl)) { if (!wEl.TryGetInt32(out w)) w = 0; }
        if (pars.TryGetProperty("h", out var hEl)) { if (!hEl.TryGetInt32(out h)) h = 0; }
        bool crop = w > 0 && h > 0;

        this.BeginInvoke(async () =>
        {
            try
            {
                if (_webView?.CoreWebView2 == null) { NotifyToast("WebView2 未就绪，无法导出"); return; }

                // #379 整页截图：导出前临时解除仪表盘滚动锁定，使完整内容（含被 overflow:auto 裁掉的底部/右侧图表）
                // 进入文档流；再用 CDP Page.captureScreenshot(captureBeyondViewport:true) 截取整页，由 clip 直接裁出图表区。
                // 彻底修复 CapturePreviewAsync 仅截当前视口、导致滚动区域外图表被截断的问题。
                Bitmap? bmp = null;
                bool unlocked = false;
                try
                {
                    await _webView.CoreWebView2.ExecuteScriptAsync("window.__unlockDash&&window.__unlockDash()");
                    unlocked = true;
                    await Task.Delay(100); // 等待布局/重绘稳定（ECharts 等可能随容器尺寸调整）

                    double dpr = 1.0;
                    try
                    {
                        var dprStr = await _webView.CoreWebView2.ExecuteScriptAsync("window.devicePixelRatio||1");
                        if (!string.IsNullOrEmpty(dprStr))
                            double.TryParse(dprStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out dpr);
                    }
                    catch { }

                    string clipJson = "{\"x\":" + Math.Max(0, x) + ",\"y\":" + Math.Max(0, y) + ",\"width\":" + Math.Max(1, w) + ",\"height\":" + Math.Max(1, h) + ",\"scale\":" + dpr.ToString(System.Globalization.CultureInfo.InvariantCulture) + "}";
                    string shotParams = "{\"format\":\"png\",\"captureBeyondViewport\":true,\"fromSurface\":true,\"clip\":" + clipJson + "}";
                    string shotJson = await _webView.CoreWebView2.CallDevToolsProtocolMethodAsync("Page.captureScreenshot", shotParams);
                    using var jdoc = JsonDocument.Parse(shotJson);
                    string b64 = jdoc.RootElement.GetProperty("data").GetString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(b64))
                    {
                        byte[] png = Convert.FromBase64String(b64);
                        using var ms2 = new MemoryStream(png);
                        using var tmp = new Bitmap(ms2);
                        bmp = new Bitmap(tmp); // 脱离流，避免流生命周期问题
                    }
                }
                catch (Exception cdpEx)
                {
                    AddInLog.Write("ExportDashboard.CDP.Error", cdpEx.ToString());
                    bmp = null;
                }
                finally
                {
                    if (unlocked)
                    {
                        try { await _webView.CoreWebView2.ExecuteScriptAsync("window.__relockDash&&window.__relockDash()"); }
                        catch { }
                    }
                }

                if (bmp == null)
                {
                    // 回退：CapturePreviewAsync 视口截图 + Bitmap 裁剪（旧逻辑；内容超出视口时仍可能截断，但保证可用）
                    using var ms = new MemoryStream();
                    await _webView.CoreWebView2.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, ms);
                    ms.Position = 0;
                    using var full = new Bitmap(ms);
                    System.Drawing.Rectangle src = new System.Drawing.Rectangle(0, 0, full.Width, full.Height);
                    if (crop)
                    {
                        double sx = full.Width / (double)Math.Max(1, _webView.ClientSize.Width);
                        double sy = full.Height / (double)Math.Max(1, _webView.ClientSize.Height);
                        int cx = (int)Math.Round(x * sx);
                        int cy = (int)Math.Round(y * sy);
                        int cw = (int)Math.Round(w * sx);
                        int ch = (int)Math.Round(h * sy);
                        cx = Math.Max(0, Math.Min(cx, full.Width - 1));
                        cy = Math.Max(0, Math.Min(cy, full.Height - 1));
                        cw = Math.Max(1, Math.Min(cw, full.Width - cx));
                        ch = Math.Max(1, Math.Min(ch, full.Height - cy));
                        src = new System.Drawing.Rectangle(cx, cy, cw, ch);
                    }
                    bmp = full.Clone(src, full.PixelFormat);
                }

                if (clipboard)
                {
                    Clipboard.SetImage(bmp);
                    NotifyToast("已复制图表到剪贴板");
                }
                else
                {
                    var dlg = new SaveFileDialog
                    {
                        Filter = "PNG 图片 (*.png)|*.png",
                        FileName = "仪表盘导出.png",
                        Title = "导出仪表盘"
                    };
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        bmp.Save(dlg.FileName, System.Drawing.Imaging.ImageFormat.Png);
                        NotifyToast("已导出：" + dlg.FileName);
                    }
                }
                bmp?.Dispose();
            }
            catch (Exception ex)
            {
                NotifyToast("导出失败：" + ex.Message);
            }
        });
    }

    /// <summary>
    /// 向 JS 推送一条轻量通知（非请求/响应），JS 端 onBridgeMessage 识别 type==='toast' 后提示。
    /// </summary>
    private void NotifyToast(string msg)
    {
        try
        {
            var json = JsonSerializer.Serialize(new { type = "toast", message = msg }, JsonOpts);
            _webView?.CoreWebView2?.PostWebMessageAsJson(json);
        }
        catch (Exception ex)
        {
            AddInLog.Write("HtmlMainForm.NotifyToast.Error", ex.ToString());
        }
    }

    private object Dispatch(string method, JsonElement pars)
    {
        switch (method)
        {
            case "getAppInfo": return GetAppInfo();
            case "listTables": return ListTables();
            case "openTable": return OpenTable(pars);
            case "saveConfig": return SaveConfig(pars);
            case "getEmbeddedConfig": return GetEmbeddedConfig();
            case "setConfigSheetVisible": return SetConfigSheetVisible(pars);
            case "saveWorkbookFile": return SaveWorkbookFile();
            case "updateCell": return UpdateCell(pars);
            case "addRow": return AddRow(pars);
            case "insertRow": return InsertRow(pars);
            case "deleteRow": return DeleteRow(pars);
            case "selectRow": return SelectRow(pars);
            case "setTopMost": return SetTopMost(pars);
            case "addField": return AddField(pars);
            case "setOpenMaximized": return SetOpenMaximized(pars);
            case "pickFolder": return PickFolder(pars);
            case "listImageFiles": return ListImageFiles(pars);
            case "getImageBase64": return GetImageBase64(pars);
            case "setVisualEditDirty": SetVisualEditDirty(pars); return new { ok = true };
            case "resolveVisualEditClose": return ResolveVisualEditClose(pars);
            case "exportDashboard": ExportDashboard(pars); return new { ok = true };
            default: return new { error = "unknown method: " + method };
        }
    }

    /// <summary>#402 接收 JS 推送的“未保存可视化调整”脏标记（甘特图手动保存模式），供关闭软件前拦截确认。</summary>
    private void SetVisualEditDirty(JsonElement pars)
    {
        try
        {
            _visualEditDirty = pars.TryGetProperty("dirty", out var d) && d.ValueKind == JsonValueKind.True;
            if (pars.TryGetProperty("count", out var c) && c.ValueKind == JsonValueKind.Number)
                _visualEditCount = c.GetInt32();
            AddInLog.Write("HtmlMainForm.VisualEditDirty", $"dirty={_visualEditDirty} count={_visualEditCount}");
        }
        catch { }
    }

    /// <summary>#406 关闭软件时若有未保存的甘特调整，改用软件自身的 HTML 弹窗（非 WinForms MessageBox，避免 VBA 风格），
    /// 由 JS 端 showConfirm3 呈现“保存 / 不保存 / 取消”，用户选择后通过 api.call('resolveVisualEditClose') 回传。</summary>
    private async Task HandlePendingVisualEditsOnClose()
    {
        try
        {
            if (_webView?.CoreWebView2 != null)
                await _webView.CoreWebView2.ExecuteScriptAsync("window.__mtPromptVisualEditClose && window.__mtPromptVisualEditClose()");
        }
        catch
        {
            // 兜底：弹窗不可用则直接强制关闭
            _forceClosing = true;
            try { Close(); } catch { }
        }
    }

    /// <summary>#406 接收 JS 端关闭确认弹窗的回传结果：save=先写回再关闭；discard=放弃再关闭；cancel=保持窗口。</summary>
    private object ResolveVisualEditClose(JsonElement pars)
    {
        try
        {
            var choice = pars.TryGetProperty("choice", out var cEl) ? (cEl.ValueKind == JsonValueKind.String ? cEl.GetString() : null) : null;
            if (choice == "cancel") return new { ok = true }; // 取消关闭，保持窗口（FormClosing 已 e.Cancel=true）
            if (choice == "discard")
            {
                try { if (_webView?.CoreWebView2 != null) _webView.CoreWebView2.ExecuteScriptAsync("window.__mtDiscardVisualEdits && window.__mtDiscardVisualEdits()"); } catch { }
            }
            else // save（默认）：先写回内存中的调整
            {
                try { if (_webView?.CoreWebView2 != null) _webView.CoreWebView2.ExecuteScriptAsync("window.__mtFlushVisualEdits && window.__mtFlushVisualEdits()"); } catch { }
            }
            _forceClosing = true;
            try { Close(); } catch { }
        }
        catch
        {
            _forceClosing = true;
            try { Close(); } catch { }
        }
        return new { ok = true };
    }

    // ───────────────────────── API 实现 ─────────────────────────

    private static object GetAppInfo()
    {
        string dll = GetDllPath();
        string modified = File.Exists(dll) ? File.GetLastWriteTime(dll).ToString("yyyy-MM-dd HH:mm:ss") : "未知";
        string host = "未知";
        try { host = HostEnvironment.GetHostDisplayName(); } catch { }
        return new { version = AppVersion.DisplayText, dllModified = modified, dllPath = dll, host, excelVersion = ExcelDnaUtil.ExcelVersion };
    }

    private static object ListTables()
    {
        var adapter = new ExcelAdapter();
        var list = adapter.GetTableSources();
        var tables = list.ConvertAll(t => new { sheet = t.SheetName, table = t.TableName, rows = t.RowCount, cols = t.ColumnCount });
        // #428-3 记忆最后使用的超级表，便于下次打开自动恢复
        string lastTable = string.Empty;
        try
        {
            string wb = adapter.GetActiveWorkbookPath();
            if (!string.IsNullOrEmpty(wb)) lastTable = ViewConfigManager.GetLastTableName(wb);
        }
        catch { }
        return new { tables, lastTable };
    }

    private object PickFolder(JsonElement pars)
    {
        string initial = "";
        try { if (pars.TryGetProperty("folder", out var fEl)) initial = fEl.GetString() ?? ""; } catch { }
        string path = "";
        using (var dlg = new FolderBrowserDialog())
        {
            dlg.Description = "选择图片文件夹";
            dlg.ShowNewFolderButton = true;
            if (Directory.Exists(initial)) dlg.SelectedPath = initial;
            if (dlg.ShowDialog(this) == DialogResult.OK) path = dlg.SelectedPath;
        }
        return new { folder = path };
    }

    private static object ListImageFiles(JsonElement pars)
    {
        string folder = "";
        try { if (pars.TryGetProperty("folder", out var fEl)) folder = fEl.GetString() ?? ""; } catch { }
        var list = new List<object>();
        if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
        {
            var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg" };
            foreach (var f in Directory.GetFiles(folder))
            {
                if (exts.Contains(Path.GetExtension(f)))
                    list.Add(new { name = Path.GetFileName(f), path = f });
            }
        }
        return new { files = list.ToArray() };
    }

    // 将本地图片读为 base64 data URI，供 HTML 端 <img> 直接显示，
    // 绕过 WebView2 在 null-origin 页面下对 file:// 图片的拦截（画册视图图片不显示的根因）
    private static object GetImageBase64(JsonElement pars)
    {
        string path = "";
        try { if (pars.TryGetProperty("path", out var pEl)) path = pEl.GetString() ?? ""; } catch { }
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return new { dataUri = "", ok = false };
        try
        {
            byte[] bytes = File.ReadAllBytes(path);
            string ext = Path.GetExtension(path).ToLowerInvariant();
            string mime = ext switch
            {
                ".png" => "image/png",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".webp" => "image/webp",
                ".svg" => "image/svg+xml",
                _ => "application/octet-stream"
            };
            return new { dataUri = "data:" + mime + ";base64," + Convert.ToBase64String(bytes), ok = true };
        }
        catch { return new { dataUri = "", ok = false }; }
    }

    private object OpenTable(JsonElement pars)
    {
        string sheet = pars.GetProperty("sheet").GetString() ?? "";
        string table = pars.GetProperty("table").GetString() ?? "";
        var adapter = new ExcelAdapter();
        var dt = adapter.ReadListObject(sheet, table);
        var mgr = new ViewConfigManager();
        string wb = adapter.GetActiveWorkbookPath();
        var config = mgr.Load(wb, dt.TableName);
        if (config.Views == null || config.Views.Count == 0)
            config = mgr.CreateDefaultConfig(dt);
        ViewConfigManager.SyncFields(config, dt);
        mgr.Save(wb, config, dt.TableName);

        _table = dt;
        _config = config;
        _workbookPath = wb;
        _sheetName = sheet;
        _tableName = dt.TableName;

        var tableView = config.Views.Find(v => v.ViewType == ViewType.Table) ?? config.Views[0];
        var visible = tableView?.VisibleFields ?? dt.FieldNames;
        return new
        {
            sheet = dt.SheetName,
            table = dt.TableName,
            fields = dt.Fields,
            rows = dt.Rows,
            config = config,
            visibleFields = visible,
            fieldOverrides = config.FieldOverrides ?? new List<FieldOverride>(),
            rowCount = dt.Rows.Count,
            colCount = dt.Fields.Count
        };
    }

    private object SaveConfig(JsonElement pars)
    {
        string json = pars.GetProperty("json").GetString() ?? "{}";
        string saveLocation = "both";
        if (pars.TryGetProperty("saveLocation", out var sl) && sl.ValueKind == JsonValueKind.String)
            saveLocation = sl.GetString() ?? "both";
        var config = JsonSerializer.Deserialize<ViewConfigFile>(json, JsonOpts);
        if (config == null) return new { saved = false };
        config.SourceFile = Path.GetFileName(_workbookPath);
        var mgr = new ViewConfigManager();
        mgr.Save(_workbookPath, config, _tableName, saveLocation);
        _config = config;
        return new { saved = true };
    }

    /// <summary>
    /// #466 读取嵌入的 _MultiTableConfig 隐藏工作表内容，供「查看配置结构与数据」对话框展示。
    /// 该表为 xlSheetVeryHidden，Excel 界面的「取消隐藏」列表不会显示它，只能由程序读取。
    /// </summary>
    private object GetEmbeddedConfig()
    {
        try
        {
            var adapter = new ExcelAdapter();
            var info = adapter.GetConfigSheetInfo();
            string filePath = string.IsNullOrEmpty(_workbookPath)
                ? string.Empty
                : ViewConfigManager.GetConfigFilePath(_workbookPath);
            if (!info.Exists)
                return new { exists = false, visible = false, json = "", length = 0, tableCount = 0, lastTable = "", filePath };
            string cfgJson = adapter.ReadConfigSheet() ?? string.Empty;
            int tableCount = 0; string lastTable = string.Empty;
            try
            {
                var wbCfg = JsonSerializer.Deserialize<WorkbookConfigFile>(cfgJson, JsonOpts);
                if (wbCfg != null)
                {
                    tableCount = wbCfg.Tables?.Count ?? 0;
                    lastTable = wbCfg.LastTableName ?? string.Empty;
                }
            }
            catch { }
            return new { exists = true, visible = info.Visible, json = cfgJson, length = cfgJson.Length, tableCount, lastTable, filePath };
        }
        catch (Exception ex)
        {
            AddInLog.Write("HtmlMainForm.GetEmbeddedConfig.Error", ex.ToString());
            return new { exists = false, visible = false, json = "", length = 0, tableCount = 0, lastTable = "", filePath = "" };
        }
    }

    /// <summary>#466 临时显示 / 恢复深度隐藏 配置工作表</summary>
    private object SetConfigSheetVisible(JsonElement pars)
    {
        bool visible = false;
        if (pars.TryGetProperty("visible", out var v)) visible = v.ValueKind == JsonValueKind.True;
        bool ok = new ExcelAdapter().SetConfigSheetVisible(visible);
        return new { ok, visible };
    }

    /// <summary>#466 保存工作簿，使嵌入配置真正落盘到 .xlsx</summary>
    private object SaveWorkbookFile()
    {
        bool saved = new ExcelAdapter().SaveWorkbookFile();
        return new { saved };
    }

    private object UpdateCell(JsonElement pars)
    {
        int rowIndex = pars.GetProperty("rowIndex").GetInt32();
        string field = pars.GetProperty("field").GetString() ?? "";
        JsonElement raw = pars.GetProperty("value");
        var fs = _table?.FindField(field);
        object? value = CoerceValue(fs?.Type ?? FieldType.Text, raw);
        var adapter = new ExcelAdapter();
        adapter.UpdateCell(_sheetName, _tableName, rowIndex, field, value);
        _table?.FindRow(rowIndex)?.SetValue(field, value);
        return new { ok = true };
    }

    private object AddRow(JsonElement pars)
    {
        var values = new Dictionary<string, object?>();
        if (pars.TryGetProperty("values", out var vEl) && vEl.ValueKind == JsonValueKind.Object)
        {
            foreach (var prop in vEl.EnumerateObject())
            {
                var fs = _table?.FindField(prop.Name);
                values[prop.Name] = CoerceValue(fs?.Type ?? FieldType.Text, prop.Value);
            }
        }
        var adapter = new ExcelAdapter();
        int newIdx = adapter.AddRow(_sheetName, _tableName, values);
        return new { rowIndex = newIdx };
    }

    private object InsertRow(JsonElement pars)
    {
        int rowIndex = pars.GetProperty("rowIndex").GetInt32();
        bool after = pars.TryGetProperty("after", out var aEl) && aEl.GetBoolean();
        var adapter = new ExcelAdapter();
        int newIdx = adapter.InsertRow(_sheetName, _tableName, rowIndex, after);
        return new { rowIndex = newIdx };
    }

    private object DeleteRow(JsonElement pars)
    {
        int rowIndex = pars.GetProperty("rowIndex").GetInt32();
        var adapter = new ExcelAdapter();
        adapter.DeleteRow(_sheetName, _tableName, rowIndex);
        return new { ok = true };
    }

    private object SelectRow(JsonElement pars)
    {
        int rowIndex = pars.GetProperty("rowIndex").GetInt32();
        var adapter = new ExcelAdapter();
        adapter.SelectRow(_sheetName, _tableName, rowIndex);
        return new { ok = true };
    }

    /// <summary>
    /// 置顶开关：默认不强制置顶；HTML 右上角“置顶”按钮可切换。
    /// 必须在 WinForms UI 线程上设置 TopMost，否则跨线程访问控件会抛异常。
    /// </summary>
    /// <summary>
    /// 置顶开关：默认不强制置顶；HTML 右上角“置顶”按钮可切换。
    /// 置顶“仅限于 Excel 窗口有效”——通过前台窗口钩子实现：仅当 Excel（或本窗口）为前台时
    /// 才 TopMost=true，一旦切换到其它软件立即取消置顶，不会盖住浏览器等其它程序。
    /// 必须在 WinForms UI 线程上访问 TopMost，否则跨线程访问控件会抛异常。
    /// </summary>
    private object SetTopMost(JsonElement pars)
    {
        bool v = pars.TryGetProperty("value", out var vEl) && vEl.ValueKind == JsonValueKind.True;
        _topMostWanted = v;
        PostToUi(() =>
        {
            try
            {
                if (v)
                {
                    uint fg; GetWindowThreadProcessId(GetForegroundWindow(), out fg);
                    TopMost = (fg == _excelPid);
                }
                else TopMost = false;
            }
            catch { }
        });
        return new { ok = true, topMost = v };
    }

    /// <summary>新增一个字段：在 Excel 超级表末尾追加一列，并将字段类型写入配置覆盖，使其类型在界面生效。</summary>
    private object AddField(JsonElement pars)
    {
        string name = pars.TryGetProperty("name", out var nEl) ? (nEl.GetString() ?? "") : "";
        string type = pars.TryGetProperty("type", out var tEl) ? (tEl.GetString() ?? "Text") : "Text";
        if (string.IsNullOrWhiteSpace(name))
            return new { ok = false, error = "字段名不能为空" };
        if (_sheetName == string.Empty || _tableName == string.Empty)
            return new { ok = false, error = "尚未打开数据源" };

        try
        {
            var adapter = new ExcelAdapter();
            adapter.AddColumn(_sheetName, _tableName, name);
        }
        catch (Exception ex)
        {
            return new { ok = false, error = "Excel 添加列失败：" + ex.Message };
        }

        // 将字段类型写入配置覆盖
        try
        {
            var mgr = new ViewConfigManager();
            var config = mgr.Load(_workbookPath, _tableName);
            config.FieldOverrides ??= new List<FieldOverride>();
            var ov = config.FieldOverrides.Find(x => x.Name == name);
            if (ov == null) { ov = new FieldOverride { Name = name }; config.FieldOverrides.Add(ov); }
            if (Enum.TryParse<FieldType>(type, out var ft)) ov.Type = ft;
            ov.UserDefined = true;
            mgr.Save(_workbookPath, config, _tableName);
        }
        catch (Exception ex)
        {
            return new { ok = false, error = "配置写入失败：" + ex.Message };
        }

        return new { ok = true };
    }

    /// <summary>根据设置决定是否在打开时最大化窗口</summary>
    private object SetOpenMaximized(JsonElement pars)
    {
        bool v = pars.TryGetProperty("value", out var vEl) && vEl.ValueKind == JsonValueKind.True;
        PostToUi(() =>
        {
            try { WindowState = v ? FormWindowState.Maximized : FormWindowState.Normal; }
            catch { }
        });
        return new { ok = true, openMaximized = v };
    }

    private static object? CoerceValue(FieldType type, JsonElement raw)
    {
        if (raw.ValueKind == JsonValueKind.Null) return null;
        switch (type)
        {
            case FieldType.Checkbox:
                return raw.ValueKind == JsonValueKind.True ||
                       (raw.ValueKind == JsonValueKind.String && string.Equals(raw.GetString(), "true", StringComparison.OrdinalIgnoreCase));
            case FieldType.Integer:
                if (raw.ValueKind == JsonValueKind.Number) return (int)Math.Round(raw.GetDouble());
                if (raw.ValueKind == JsonValueKind.String && int.TryParse(raw.GetString(), out int i)) return i;
                return 0;
            case FieldType.Number:
            case FieldType.Currency:
            case FieldType.Percentage:
                if (raw.ValueKind == JsonValueKind.Number) return raw.GetDouble();
                if (raw.ValueKind == JsonValueKind.String && double.TryParse(raw.GetString(), out double d)) return d;
                return 0d;
            case FieldType.Date:
            case FieldType.DateTime:
                if (raw.ValueKind == JsonValueKind.String && DateTime.TryParse(raw.GetString(), out var dt))
                    return dt == default ? null : dt;
                return null;
            default:
                if (raw.ValueKind == JsonValueKind.String) return raw.GetString();
                if (raw.ValueKind == JsonValueKind.Number) return raw.GetDouble().ToString();
                if (raw.ValueKind == JsonValueKind.True) return "true";
                if (raw.ValueKind == JsonValueKind.False) return "false";
                return raw.GetRawText();
        }
    }

    // ───────────────────────── Ribbon 入口 ─────────────────────────

    public void Reload()
    {
        if (_coreReady)
            _ = _webView.CoreWebView2.ExecuteScriptAsync("window.__mtReload && window.__mtReload()");
    }

    public void SaveConfigNow()
    {
        if (_coreReady)
            _ = _webView.CoreWebView2.ExecuteScriptAsync("window.__mtSave && window.__mtSave()");
    }
}
