using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using FormsTimer = System.Windows.Forms.Timer;

namespace MultiTableAddin.TaskPane;

public interface ITaskPaneHostControl
{
}

internal interface ITaskPaneCloseAware
{
    void PrepareForClose(TaskPaneCloseContext context);
}

internal static class TaskPaneHostFactory
{
    internal const int DefaultTaskPaneWidth = 1000;
    internal static Guid CtpAddInClsId => new(WpsCompatSettings.CtpAddInGuid);

    internal static Control CreateHostedView(
        Func<System.Windows.UIElement> viewFactory,
        string logKey,
        string fallbackMessage)
    {
        try
        {
            return new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = viewFactory()
            };
        }
        catch (Exception ex)
        {
            AddInLog.Write(logKey, ex.ToString());
            return CreateFallbackControl(fallbackMessage, ex);
        }
    }

    internal static Control CreateFallbackControl(string message, Exception ex)
    {
        return new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Text = message + Environment.NewLine + Environment.NewLine + ex
        };
    }

    internal static void ApplyPreferredWidth(CustomTaskPane taskPane, int width, string logKey)
    {
        try
        {
            taskPane.Width = width;
            AddInLog.Write(logKey + ".Width", "AppliedImmediately=" + width);
        }
        catch (Exception ex)
        {
            AddInLog.Write(logKey + ".Width.Error", "ApplyImmediately=" + width + " | " + ex);
        }
    }

    internal static void ScheduleWpsWidthRetry(CustomTaskPane taskPane, int width, string logKey)
    {
        if (!HostEnvironment.IsWpsEt())
        {
            return;
        }

        FormsTimer retryTimer = new FormsTimer
        {
            Interval = 500
        };

        retryTimer.Tick += (_, _) =>
        {
            retryTimer.Stop();
            retryTimer.Dispose();

            try
            {
                taskPane.Width = width;
                AddInLog.Write(logKey + ".Width", "AppliedAfterDelay=" + width);
            }
            catch (Exception ex)
            {
                AddInLog.Write(logKey + ".Width.Error", "ApplyAfterDelay=" + width + " | " + ex);
            }
        };

        AddInLog.Write(logKey + ".Width", "ScheduledRetryMs=500; TargetWidth=" + width);
        retryTimer.Start();
    }
}

internal sealed class TaskPaneStatusSnapshot
{
    internal TaskPaneStatusSnapshot(
        string paneName,
        string activeWindowKey,
        string activeWindowCaption,
        int managedWindowCount,
        bool hasActiveWindowPane,
        bool isActiveWindowPaneVisible)
    {
        PaneName = paneName;
        ActiveWindowKey = activeWindowKey;
        ActiveWindowCaption = activeWindowCaption;
        ManagedWindowCount = managedWindowCount;
        HasActiveWindowPane = hasActiveWindowPane;
        IsActiveWindowPaneVisible = isActiveWindowPaneVisible;
    }

    internal string PaneName { get; }

    internal string ActiveWindowKey { get; }

    internal string ActiveWindowCaption { get; }

    internal int ManagedWindowCount { get; }

    internal bool HasActiveWindowPane { get; }

    internal bool IsActiveWindowPaneVisible { get; }

    internal string VisibilityText => !HasActiveWindowPane
        ? "当前窗口尚未创建窗格"
        : (IsActiveWindowPaneVisible ? "当前窗口窗格可见" : "当前窗口窗格已创建但隐藏");

    internal string SummaryText => string.Format(
        CultureInfo.InvariantCulture,
        "窗格类型：{0}\r\n活动窗口键：{1}\r\n活动窗口标题：{2}\r\n已管理窗口数：{3}\r\n状态：{4}",
        PaneName,
        ActiveWindowKey,
        ActiveWindowCaption,
        ManagedWindowCount,
        VisibilityText);
}

internal sealed class ActiveTaskPaneWindow
{
    internal ActiveTaskPaneWindow(object parentWindow, string windowKey, string windowCaption)
    {
        ParentWindow = parentWindow;
        WindowKey = windowKey;
        WindowCaption = windowCaption;
    }

    internal object ParentWindow { get; }

    internal string WindowKey { get; }

    internal string WindowCaption { get; }
}

internal sealed class ManagedTaskPaneEntry
{
    internal ManagedTaskPaneEntry(
        string windowKey,
        string windowCaption,
        CustomTaskPane taskPane,
        CustomTaskPaneEvents_VisibleStateChangeEventHandler visibleStateHandler)
    {
        WindowKey = windowKey;
        WindowCaption = windowCaption;
        TaskPane = taskPane;
        VisibleStateHandler = visibleStateHandler;
        DesiredVisible = true;
    }

    internal string WindowKey { get; }

    internal string WindowCaption { get; set; }

    internal CustomTaskPane TaskPane { get; }

    internal CustomTaskPaneEvents_VisibleStateChangeEventHandler VisibleStateHandler { get; }

    internal bool DesiredVisible { get; set; }
}

internal sealed class ManagedCustomTaskPaneCollection
{
    private readonly object _syncRoot = new();
    private readonly Dictionary<string, ManagedTaskPaneEntry> _entries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Type _hostControlType;
    private readonly string _paneTitle;
    private readonly string _paneName;
    private readonly string _logKey;

    internal ManagedCustomTaskPaneCollection(Type hostControlType, string paneTitle, string paneName, string logKey)
    {
        _hostControlType = hostControlType;
        _paneTitle = paneTitle;
        _paneName = paneName;
        _logKey = logKey;
    }

    internal TaskPaneStatusSnapshot ShowForActiveWindow(string source)
    {
        lock (_syncRoot)
        {
            PruneStaleEntries(source);
            if (!TryGetActiveWindow(out ActiveTaskPaneWindow? activeWindow) || activeWindow == null)
            {
                AddInLog.Write(_logKey + ".NoActiveWindow", source);
                return CreateSnapshot("无活动窗口", "无活动窗口", false, false);
            }

            ManagedTaskPaneEntry entry = GetOrCreateEntry(activeWindow, source);
            entry.DesiredVisible = true;
            entry.WindowCaption = activeWindow.WindowCaption;
            entry.TaskPane.Visible = true;
            TaskPaneHostFactory.ApplyPreferredWidth(entry.TaskPane, TaskPaneHostFactory.DefaultTaskPaneWidth, _logKey);
            TaskPaneHostFactory.ScheduleWpsWidthRetry(entry.TaskPane, TaskPaneHostFactory.DefaultTaskPaneWidth, _logKey);
            AddInLog.Write(_logKey + ".Visible", source + "; WindowKey=" + activeWindow.WindowKey);
            return CreateSnapshot(activeWindow.WindowKey, activeWindow.WindowCaption, true, true);
        }
    }

    internal TaskPaneStatusSnapshot GetSnapshot()
    {
        lock (_syncRoot)
        {
            PruneStaleEntries("Snapshot");
            if (!TryGetActiveWindow(out ActiveTaskPaneWindow? activeWindow) || activeWindow == null)
            {
                return CreateSnapshot("无活动窗口", "无活动窗口", false, false);
            }

            bool hasPane = _entries.TryGetValue(activeWindow.WindowKey, out ManagedTaskPaneEntry? entry);
            bool isVisible = hasPane && entry != null && SafeGetVisible(entry.TaskPane);
            return CreateSnapshot(activeWindow.WindowKey, activeWindow.WindowCaption, hasPane, isVisible);
        }
    }

    internal void DisposeAll(string source)
    {
        lock (_syncRoot)
        {
            foreach (string windowKey in new List<string>(_entries.Keys))
            {
                RemoveEntry(windowKey, source, "DisposeAll");
            }
        }
    }

    private ManagedTaskPaneEntry GetOrCreateEntry(ActiveTaskPaneWindow activeWindow, string source)
    {
        if (_entries.TryGetValue(activeWindow.WindowKey, out ManagedTaskPaneEntry? existingEntry))
        {
            return existingEntry;
        }

        CustomTaskPane? taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(
            _hostControlType,
            _paneTitle,
            activeWindow.ParentWindow);
        if (taskPane == null)
        {
            throw new InvalidOperationException("CreateCustomTaskPane 返回 null。");
        }

        taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        taskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
        TaskPaneHostFactory.ApplyPreferredWidth(taskPane, TaskPaneHostFactory.DefaultTaskPaneWidth, _logKey);

        CustomTaskPaneEvents_VisibleStateChangeEventHandler handler = pane => OnVisibleStateChanged(activeWindow.WindowKey, pane);
        taskPane.VisibleStateChange += handler;

        ManagedTaskPaneEntry entry = new ManagedTaskPaneEntry(activeWindow.WindowKey, activeWindow.WindowCaption, taskPane, handler);
        _entries[activeWindow.WindowKey] = entry;
        AddInLog.Write(_logKey + ".Created", source + "; WindowKey=" + activeWindow.WindowKey + "; WindowCaption=" + activeWindow.WindowCaption);
        return entry;
    }

    private void OnVisibleStateChanged(string windowKey, CustomTaskPane taskPane)
    {
        lock (_syncRoot)
        {
            if (_entries.TryGetValue(windowKey, out ManagedTaskPaneEntry? entry))
            {
                entry.DesiredVisible = SafeGetVisible(taskPane);
                AddInLog.Write(_logKey + ".VisibleStateChanged", "WindowKey=" + windowKey + "; Visible=" + entry.DesiredVisible);
            }
        }
    }

    private void PruneStaleEntries(string source)
    {
        foreach (KeyValuePair<string, ManagedTaskPaneEntry> pair in new List<KeyValuePair<string, ManagedTaskPaneEntry>>(_entries))
        {
            if (!CanAccessTaskPane(pair.Value.TaskPane))
            {
                RemoveEntry(pair.Key, source, "Stale");
            }
        }
    }

    private void RemoveEntry(string windowKey, string source, string reason)
    {
        if (!_entries.TryGetValue(windowKey, out ManagedTaskPaneEntry? entry))
        {
            return;
        }

        try
        {
            entry.TaskPane.VisibleStateChange -= entry.VisibleStateHandler;
        }
        catch
        {
        }

        try
        {
            entry.TaskPane.Delete();
        }
        catch
        {
        }

        _entries.Remove(windowKey);
        AddInLog.Write(_logKey + ".Removed", source + "; WindowKey=" + windowKey + "; Reason=" + reason);
    }

    private TaskPaneStatusSnapshot CreateSnapshot(string windowKey, string windowCaption, bool hasPane, bool isVisible)
    {
        return new TaskPaneStatusSnapshot(
            _paneName,
            windowKey,
            windowCaption,
            _entries.Count,
            hasPane,
            isVisible);
    }

    private static bool TryGetActiveWindow(out ActiveTaskPaneWindow? activeWindow)
    {
        activeWindow = null;
        try
        {
            dynamic application = ExcelDnaUtil.Application;
            object? window = application?.ActiveWindow;
            if (window == null)
            {
                return false;
            }

            string windowKey = TryGetWindowKey(window);
            string windowCaption = TryGetWindowCaption(window);
            activeWindow = new ActiveTaskPaneWindow(window, windowKey, windowCaption);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string TryGetWindowKey(object window)
    {
        try
        {
            string hwnd = Convert.ToString(((dynamic)window).Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(hwnd))
            {
                return "HWND=" + hwnd.Trim();
            }
        }
        catch
        {
        }

        try
        {
            string caption = Convert.ToString(((dynamic)window).Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            string windowNumber = Convert.ToString(((dynamic)window).WindowNumber, CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(caption) || !string.IsNullOrWhiteSpace(windowNumber))
            {
                return "Caption=" + caption.Trim() + ";WindowNumber=" + windowNumber.Trim();
            }
        }
        catch
        {
        }

        return "Window#" + RuntimeHelpers.GetHashCode(window).ToString(CultureInfo.InvariantCulture);
    }

    private static string TryGetWindowCaption(object window)
    {
        try
        {
            string caption = Convert.ToString(((dynamic)window).Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(caption))
            {
                return caption.Trim();
            }
        }
        catch
        {
        }

        return "未命名窗口";
    }

    private static bool CanAccessTaskPane(CustomTaskPane taskPane)
    {
        try
        {
            _ = taskPane.Visible;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool SafeGetVisible(CustomTaskPane taskPane)
    {
        try
        {
            return taskPane.Visible;
        }
        catch
        {
            return false;
        }
    }

}

[ComVisible(true)]
[ComDefaultInterface(typeof(ITaskPaneHostControl))]
public class TaskPaneHostControl : UserControl, ITaskPaneHostControl
{
    public TaskPaneHostControl()
    {
        Dock = DockStyle.Fill;
        Controls.Add(TaskPaneHostFactory.CreateHostedView(
            () => new TaskPaneWpfView(),
            "CTP.HostControlFallback",
            "WPF 窗格初始化失败，请检查依赖文件是否已复制到输出目录。"));
    }
}

internal static class TaskPaneManager
{
    private static readonly ManagedCustomTaskPaneCollection Collection = new(
        typeof(TaskPaneHostControl),
        WpsCompatSettings.TaskPaneTitle,
        "主任务窗格",
        "CTP");

    internal static TaskPaneStatusSnapshot Show(string source)
    {
        return Collection.ShowForActiveWindow(source);
    }

    internal static TaskPaneStatusSnapshot GetStatusSnapshot()
    {
        return Collection.GetSnapshot();
    }

    internal static void DisposeAll(string source)
    {
        Collection.DisposeAll(source);
    }
}

[ComVisible(true)]
[ComDefaultInterface(typeof(ITaskPaneHostControl))]
public class WebTaskPaneHostControl : UserControl, ITaskPaneHostControl
{
    private readonly ElementHost? _elementHost;
    private readonly WebTaskPaneWpfView? _view;

    public WebTaskPaneHostControl()
    {
        Dock = DockStyle.Fill;
        try
        {
            _view = new WebTaskPaneWpfView();
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _view
            };
            Controls.Add(_elementHost);
        }
        catch (Exception ex)
        {
            AddInLog.Write("WebView2.CTP.HostControlFallback", ex.ToString());
            Controls.Add(TaskPaneHostFactory.CreateFallbackControl(
                "WPF WebView2 窗格初始化失败，请确认依赖文件与 WebView2 Runtime 已准备完整。",
                ex));
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            TaskPaneCloseContext context = AddInLifecycle.CreateTaskPaneCloseContext(nameof(WebTaskPaneHostControl));
            try
            {
                (_view as ITaskPaneCloseAware)?.PrepareForClose(context);
            }
            catch (Exception ex)
            {
                AddInLog.Write("WebView2.CTP.PrepareForClose.Error", ex.ToString());
            }

            try
            {
                if (_elementHost != null)
                {
                    _elementHost.Child = null;
                }
            }
            catch (Exception ex)
            {
                AddInLog.Write("WebView2.CTP.DetachChild.Error", ex.ToString());
            }
        }

        base.Dispose(disposing);
    }
}

internal static class WebTaskPaneManager
{
    private static readonly ManagedCustomTaskPaneCollection Collection = new(
        typeof(WebTaskPaneHostControl),
        WpsCompatSettings.TaskPaneTitle + " - WebView2",
        "WebView2 窗格",
        "WebView2.CTP");

    internal static TaskPaneStatusSnapshot Show(string source)
    {
        return Collection.ShowForActiveWindow(source);
    }

    internal static TaskPaneStatusSnapshot GetStatusSnapshot()
    {
        return Collection.GetSnapshot();
    }

    internal static void DisposeAll(string source)
    {
        Collection.DisposeAll(source);
    }
}
