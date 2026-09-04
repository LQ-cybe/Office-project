using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;
using System.Text;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using NLog;
using NLog.Config;
using NLog.Targets;
using MultiTableAddin.TaskPane;

namespace MultiTableAddin;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        AddInRuntime.Initialize();
        AddInLog.Write("AutoOpen");

        try
        {
            HandyControlRuntime.Initialize();
            AddInLog.Write("HandyControl.Initialize", "Success");
        }
        catch (Exception ex)
        {
            AddInLog.Write("HandyControl.Initialize.Error", ex.ToString());
        }

        try
        {
            IntelliSenseServer.Install();
            AddInLog.Write("IntelliSense.Install", "Success");
        }
        catch (Exception ex)
        {
            AddInLog.Write("IntelliSense.Install.Error", ex.ToString());
        }
    }

    public void AutoClose()
    {
        AddInLifecycle.MarkAutoClose("AutoClose");
        TaskPaneManager.DisposeAll("AutoClose");
        WebTaskPaneManager.DisposeAll("AutoClose");

        try
        {
            IntelliSenseServer.Uninstall();
            AddInLog.Write("IntelliSense.Uninstall", "Success");
        }
        catch (Exception ex)
        {
            AddInLog.Write("IntelliSense.Uninstall.Error", ex.ToString());
        }

        AddInLog.Write("AutoClose");
    }
}

internal sealed class TaskPaneCloseContext
{
    internal TaskPaneCloseContext(string source, string reason, bool avoidAggressiveDispose)
    {
        Source = source;
        Reason = reason;
        AvoidAggressiveDispose = avoidAggressiveDispose;
    }

    internal string Source { get; }

    internal string Reason { get; }

    internal bool AvoidAggressiveDispose { get; }
}

internal static class AddInLifecycle
{
    private static int _autoCloseStarted;
    private static int _processExitStarted;
    private static int _domainUnloadStarted;
    private static int _assemblyUnloadStarted;

    internal static bool IsHostShuttingDown =>
        Volatile.Read(ref _processExitStarted) == 1 ||
        Volatile.Read(ref _domainUnloadStarted) == 1 ||
        Volatile.Read(ref _assemblyUnloadStarted) == 1;

    internal static void MarkAutoClose(string source)
    {
        if (Interlocked.Exchange(ref _autoCloseStarted, 1) == 0)
        {
            AddInLog.Write("Lifecycle.AutoClose", source);
        }
    }

    internal static void MarkProcessExit(string source)
    {
        if (Interlocked.Exchange(ref _processExitStarted, 1) == 0)
        {
            AddInLog.Write("Lifecycle.ProcessExit", source);
        }
    }

    internal static void MarkDomainUnload(string source)
    {
        if (Interlocked.Exchange(ref _domainUnloadStarted, 1) == 0)
        {
            AddInLog.Write("Lifecycle.DomainUnload", source);
        }
    }

    internal static void MarkAssemblyUnload(string source)
    {
        if (Interlocked.Exchange(ref _assemblyUnloadStarted, 1) == 0)
        {
            AddInLog.Write("Lifecycle.AssemblyUnload", source);
        }
    }

    internal static TaskPaneCloseContext CreateTaskPaneCloseContext(string source)
    {
        return new TaskPaneCloseContext(source, GetCloseReason(), IsHostShuttingDown);
    }

    private static string GetCloseReason()
    {
        if (Volatile.Read(ref _processExitStarted) == 1)
        {
            return "ProcessExit";
        }

        if (Volatile.Read(ref _domainUnloadStarted) == 1)
        {
            return "DomainUnload";
        }

        if (Volatile.Read(ref _assemblyUnloadStarted) == 1)
        {
            return "AssemblyUnload";
        }

        if (Volatile.Read(ref _autoCloseStarted) == 1)
        {
            return "AutoClose";
        }

        return "Unknown";
    }
}

internal static class AddInLog
{
    private const long MaxLogFileSizeBytes = 5L * 1024L * 1024L;
    private const int LogRetentionDays = 30;
    private static int _configured;
    private static readonly object SyncRoot = new();

    internal static string LogsDirectory => Path.Combine(AddInRuntime.GetAddInDirectory(), "logs");

    internal static string LogFilePath => ResolveCurrentLogFilePath(DateTime.Now);

    internal static void Write(string stage)
    {
        Write(stage, null);
    }

    internal static void Write(string stage, string? details)
    {
        try
        {
            EnsureConfigured();

            LogEventInfo logEvent = new LogEventInfo(LogLevel.Info, "AddIn", details ?? string.Empty);
            logEvent.Properties["Stage"] = stage;
            logEvent.Properties["Host"] = HostEnvironment.GetHostDisplayName();
            logEvent.Properties["Process"] = Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty;
            logEvent.Properties["ThreadId"] = Environment.CurrentManagedThreadId;
            logEvent.Properties["ApartmentState"] = Thread.CurrentThread.GetApartmentState();

            LogManager.GetLogger("AddIn").Log(logEvent);
        }
        catch (Exception ex)
        {
            WriteFallback(stage, details, ex);
        }
    }

    private static void EnsureConfigured()
    {
        if (Volatile.Read(ref _configured) == 1)
        {
            return;
        }

        lock (SyncRoot)
        {
            if (_configured == 1)
            {
                return;
            }

            Directory.CreateDirectory(LogsDirectory);
            CleanupExpiredLogs();

            LoggingConfiguration config = new LoggingConfiguration();
            config.Variables["addinLogDir"] = LogsDirectory.Replace("\\", "/");

            FileTarget fileTarget = new FileTarget("addin-file")
            {
                FileName = "${var:addinLogDir}/${shortdate}.log",
                ArchiveFileName = "${var:addinLogDir}/${shortdate}_{#}.log",
                ArchiveAboveSize = MaxLogFileSizeBytes,
                ArchiveNumbering = ArchiveNumberingMode.Sequence,
                ConcurrentWrites = true,
                KeepFileOpen = false,
                OpenFileCacheTimeout = 5,
                AutoFlush = true,
                Encoding = new UTF8Encoding(true),
                Layout = "${longdate} | ${event-properties:item=Stage} | Host=${event-properties:item=Host} | Process=${event-properties:item=Process} | Thread=${event-properties:item=ThreadId} | Apartment=${event-properties:item=ApartmentState} | Details=${message}"
            };

            config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget);
            LogManager.Configuration = config;
            LogManager.ReconfigExistingLoggers();
            Volatile.Write(ref _configured, 1);
        }
    }

    private static string ResolveLogFilePath(DateTime timestamp, bool ensureDirectory)
    {
        if (ensureDirectory)
        {
            Directory.CreateDirectory(LogsDirectory);
        }

        string baseName = timestamp.ToString("yyyy-MM-dd");
        string candidate = Path.Combine(LogsDirectory, baseName + ".log");
        if (CanAppendToLogFile(candidate))
        {
            return candidate;
        }

        for (int index = 1; index < 1000; index++)
        {
            candidate = Path.Combine(LogsDirectory, string.Format("{0}_{1}.log", baseName, index));
            if (CanAppendToLogFile(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(LogsDirectory, baseName + "_999.log");
    }

    private static string ResolveCurrentLogFilePath(DateTime timestamp)
    {
        Directory.CreateDirectory(LogsDirectory);

        string baseName = timestamp.ToString("yyyy-MM-dd");
        string primaryPath = Path.Combine(LogsDirectory, baseName + ".log");
        if (File.Exists(primaryPath))
        {
            return primaryPath;
        }

        for (int index = 1; index < 1000; index++)
        {
            string archivedPath = Path.Combine(LogsDirectory, string.Format("{0}_{1}.log", baseName, index));
            if (File.Exists(archivedPath))
            {
                return archivedPath;
            }
        }

        return primaryPath;
    }

    private static bool CanAppendToLogFile(string filePath)
    {
        if (!File.Exists(filePath)) {
            return true;
        }

        FileInfo fileInfo = new FileInfo(filePath);
        return fileInfo.Length < MaxLogFileSizeBytes;
    }

    private static void CleanupExpiredLogs()
    {
        DateTime cutoffTime = DateTime.Now.AddDays(-LogRetentionDays);
        foreach (string filePath in Directory.EnumerateFiles(LogsDirectory, "*.log", SearchOption.TopDirectoryOnly))
        {
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                if (fileInfo.LastWriteTime < cutoffTime)
                {
                    fileInfo.Delete();
                }
            }
            catch
            {
            }
        }
    }

    private static void WriteFallback(string stage, string? details, Exception exception)
    {
        try
        {
            string filePath = ResolveLogFilePath(DateTime.Now, true);
            string line = string.Format(
                "{0:yyyy-MM-dd HH:mm:ss} | {1} | Host={2} | Process={3} | Thread={4} | Apartment={5} | Details={6} | LoggingFallback={7}",
                DateTime.Now,
                stage,
                HostEnvironment.GetHostDisplayName(),
                Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty,
                Environment.CurrentManagedThreadId,
                Thread.CurrentThread.GetApartmentState(),
                details ?? string.Empty,
                exception);

            File.AppendAllText(filePath, line + Environment.NewLine, new UTF8Encoding(true));
        }
        catch
        {
        }
    }
}

internal static class HostEnvironment
{
    internal static bool IsWpsEt()
    {
        return ExcelDnaUtil.IsET;
    }

    internal static string GetHostDisplayName()
    {
        return IsWpsEt() ? "WPS ET" : "Microsoft Excel";
    }

    internal static string GetRuntimeDisplayText()
    {
        string frameworkDescription = RuntimeInformation.FrameworkDescription;
        string targetFramework = typeof(AddIn).Assembly
            .GetCustomAttribute<System.Runtime.Versioning.TargetFrameworkAttribute>()?.FrameworkName
            ?? "unknown";
        string processBitness = Environment.Is64BitProcess ? "x64" : "x86";

        return string.Format(
            "Host={0}\r\nExcelVersion={1}\r\nFrameworkDescription={2}\r\nEnvironment.Version={3}\r\nTargetFramework={4}\r\nProcessBitness={5}\r\nBaseDirectory={6}\r\n{7}",
            GetHostDisplayName(),
            ExcelDnaUtil.ExcelVersion,
            frameworkDescription,
            Environment.Version,
            targetFramework,
            processBitness,
            AppContext.BaseDirectory,
            HandyControlRuntime.GetStatusSnapshot());
    }
}

internal static class AddInRuntime
{
    private static int _initialized;

    internal static void Initialize()
    {
        if (Interlocked.Exchange(ref _initialized, 1) == 1)
        {
            return;
        }

        AppDomain.CurrentDomain.ProcessExit += (_, _) => AddInLifecycle.MarkProcessExit("AppDomain.ProcessExit");
        AppDomain.CurrentDomain.DomainUnload += (_, _) => AddInLifecycle.MarkDomainUnload("AppDomain.DomainUnload");
        AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

        AssemblyLoadContext? loadContext = AssemblyLoadContext.GetLoadContext(typeof(AddInRuntime).Assembly);
        if (loadContext != null)
        {
            loadContext.Resolving += LoadContext_Resolving;
            loadContext.Unloading += _ => AddInLifecycle.MarkAssemblyUnload("AssemblyLoadContext.Unloading");
        }

        foreach (string assemblyName in new[]
        {
            "HandyControl",
            "NLog",
            "Microsoft.Web.WebView2.Core",
            "Microsoft.Web.WebView2.WinForms"
        })
        {
            TryLoadAssemblyFromAddInDirectory(assemblyName);
        }
    }

    internal static string GetAddInDirectory()
    {
        try
        {
            string xllPath = ExcelDnaUtil.XllPath;
            if (!string.IsNullOrWhiteSpace(xllPath))
            {
                string? xllDirectory = Path.GetDirectoryName(xllPath);
                if (!string.IsNullOrWhiteSpace(xllDirectory))
                {
                    return xllDirectory;
                }
            }
        }
        catch
        {
        }

        string? assemblyDirectory = Path.GetDirectoryName(typeof(AddInRuntime).Assembly.Location);
        if (!string.IsNullOrWhiteSpace(assemblyDirectory))
        {
            return assemblyDirectory;
        }

        return AppContext.BaseDirectory;
    }

    internal static string GetWebView2DataDirectory()
    {
        return Path.Combine(
            GetAddInDirectory(),
            "webview2",
            Environment.Is64BitProcess ? "x64" : "x86");
    }

    private static Assembly? CurrentDomain_AssemblyResolve(object? sender, ResolveEventArgs args)
    {
        return TryLoadAssemblyFromAddInDirectory(new AssemblyName(args.Name).Name);
    }

    private static Assembly? LoadContext_Resolving(AssemblyLoadContext context, AssemblyName assemblyName)
    {
        return TryLoadAssemblyFromAddInDirectory(assemblyName.Name, context);
    }

    private static Assembly? TryLoadAssemblyFromAddInDirectory(string? simpleName, AssemblyLoadContext? loadContext = null)
    {
        if (string.IsNullOrWhiteSpace(simpleName))
        {
            return null;
        }

        Assembly? loadedAssembly = AppDomain.CurrentDomain
            .GetAssemblies()
            .FirstOrDefault(assembly => string.Equals(assembly.GetName().Name, simpleName, StringComparison.OrdinalIgnoreCase));
        if (loadedAssembly != null)
        {
            return loadedAssembly;
        }

        string assemblyPath = Path.Combine(GetAddInDirectory(), simpleName + ".dll");
        if (!File.Exists(assemblyPath))
        {
            return null;
        }

        try
        {
            if (loadContext != null)
            {
                return loadContext.LoadFromAssemblyPath(assemblyPath);
            }

            AssemblyLoadContext? currentContext = AssemblyLoadContext.GetLoadContext(typeof(AddInRuntime).Assembly);
            if (currentContext != null)
            {
                return currentContext.LoadFromAssemblyPath(assemblyPath);
            }

            return Assembly.LoadFrom(assemblyPath);
        }
        catch (Exception ex)
        {
            AddInLog.Write("AssemblyResolve.Error", ex.ToString());
            return null;
        }
    }
}
