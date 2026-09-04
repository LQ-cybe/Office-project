using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using HandyControl.Controls;
using WpfApplication = System.Windows.Application;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfPanel = System.Windows.Controls.Panel;
using WpfScrollViewer = System.Windows.Controls.ScrollViewer;
using WpfWindow = System.Windows.Window;

namespace MultiTableAddin;

internal static class HandyControlRuntime
{
    internal const string DefaultAccentPresetKey = "default";
    private const string GlobalGrowlToken = "__GLOBAL_GROWL__";
    private static readonly object SyncRoot = new();
    private const string ThemeOverridesRelativePath = "Resources/ThemeOverrides.xaml";
    private const string DatePickerStylesRelativePath = "Resources/DatePickerStyles.xaml";
    private static readonly string[] HandyControlResourceUris =
    {
        "pack://application:,,,/HandyControl;component/Themes/SkinDefault.xaml",
        "pack://application:,,,/HandyControl;component/Themes/Theme.xaml"
    };
    private static readonly (string ColorKey, string BrushKey)[] AccentPaletteEntries =
    {
        ("PrimaryColor", "PrimaryBrush"),
        ("LightPrimaryColor", "LightPrimaryBrush"),
        ("DarkPrimaryColor", "DarkPrimaryBrush"),
        ("AccentColor", "AccentBrush"),
        ("DarkAccentColor", "DarkAccentBrush")
    };
    private static readonly AccentPresetOption[] AccentPresetOptions =
    {
        new(DefaultAccentPresetKey, "跟随默认", null, null, null, "恢复到 HandyControl 官方默认 Accent 或 ThemeOverrides 当前基线。"),
        new("red", "红色", "#FF3B30", "#FF8A80", "#D92D20", "适合测试高关注、高对比度的品牌主色。"),
        new("orange", "橙色", "#F59E0B", "#FCD34D", "#B45309", "适合偏活跃、运营风格或强调引导按钮的场景。"),
        new("gold", "金色", "#D4A017", "#EACD77", "#9B6B00", "适合会员、权益、数据看板类视觉方案。"),
        new("green", "绿色", "#22C55E", "#86EFAC", "#15803D", "适合成功态较多、工具效率型页面。"),
        new("blue", "蓝色", "#2563EB", "#93C5FD", "#1D4ED8", "适合企业后台、政企、常规工具型产品。"),
        new("purple", "紫色", "#8B5CF6", "#C4B5FD", "#6D28D9", "适合偏创意、品牌感较强的视觉风格。")
    };
    private static ThemePaletteSnapshot? _defaultAccentSnapshot;
    private static WpfWindow? _globalGrowlHostWindow;
    private static WpfPanel? _globalGrowlHostPanel;
    private static int _initialized;

    internal static WpfApplication CurrentApplication => EnsureApplication();

    internal static bool IsInitialized => WpfApplication.Current != null || Volatile.Read(ref _initialized) == 1;

    internal static void Initialize()
    {
        WpfApplication application = EnsureApplication();
        AddInLog.Write("HandyControl.Initialize.State", BuildStatusSnapshot(application));
    }

    internal static void Invoke(Action action)
    {
        WpfApplication application = EnsureApplication();
        Dispatcher dispatcher = application.Dispatcher;

        if (dispatcher.CheckAccess())
        {
            action();
            return;
        }

        dispatcher.Invoke(action);
    }

    internal static void BeginInvoke(Action action)
    {
        WpfApplication application = EnsureApplication();
        Dispatcher dispatcher = application.Dispatcher;

        if (dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(action, DispatcherPriority.Normal);
            return;
        }

        dispatcher.BeginInvoke(action, DispatcherPriority.Normal);
    }

    internal static bool TryInfoGlobal(string message)
    {
        return TryNotifyGlobal(message, Growl.InfoGlobal, Growl.Info, "Info");
    }

    internal static bool TrySuccessGlobal(string message)
    {
        return TryNotifyGlobal(message, Growl.SuccessGlobal, Growl.Success, "Success");
    }

    internal static bool TryWarningGlobal(string message)
    {
        return TryNotifyGlobal(message, Growl.WarningGlobal, Growl.Warning, "Warning");
    }

    internal static bool TryErrorGlobal(string message)
    {
        return TryNotifyGlobal(message, Growl.ErrorGlobal, Growl.Error, "Error");
    }

    internal static IReadOnlyList<AccentPresetOption> GetAccentPresetOptions()
    {
        return AccentPresetOptions;
    }

    internal static AccentPresetOption GetAccentPresetOrDefault(string? presetKey)
    {
        return AccentPresetOptions.FirstOrDefault(item =>
                   string.Equals(item.Key, presetKey, StringComparison.OrdinalIgnoreCase))
               ?? AccentPresetOptions[0];
    }

    internal static bool TryApplyAccentPreset(string? presetKey, out AccentPresetOption preset)
    {
        preset = GetAccentPresetOrDefault(presetKey);
        AccentPresetOption selectedPreset = preset;

        try
        {
            Invoke(() => ApplyAccentPreset(selectedPreset));
            AddInLog.Write(
                "HandyControl.AccentPreset",
                "Preset=" + selectedPreset.Key + " | " + GetStatusSnapshot());
            return true;
        }
        catch (Exception ex)
        {
            AddInLog.Write(
                "HandyControl.AccentPreset.Error",
                "Preset=" + selectedPreset.Key + " | " + ex);
            return false;
        }
    }

    private static bool TryNotifyGlobal(string message, Action<string> fallbackNotifyAction, Action<string, string> tokenNotifyAction, string level)
    {
        try
        {
            AddInLog.Write("Growl." + level + "Global.Attempt", "Message=" + message + " | " + GetStatusSnapshot());
            Invoke(() =>
            {
                if (EnsureGlobalGrowlHost())
                {
                    tokenNotifyAction(message, GlobalGrowlToken);
                    return;
                }

                fallbackNotifyAction(message);
            });
            AddInLog.Write("Growl." + level + "Global", "Message=" + message + " | " + GetStatusSnapshot());
            return true;
        }
        catch (Exception ex)
        {
            AddInLog.Write("Growl." + level + "Global.Error", "Message=" + message + " | " + GetStatusSnapshot() + " | " + ex);
            return false;
        }
    }

    internal static string GetStatusSnapshot()
    {
        WpfApplication? application = WpfApplication.Current;
        if (application == null)
        {
            return string.Format(
                "HandyControlState: InitializedFlag={0}; ApplicationCurrent=False; Thread={1}; Apartment={2}",
                Volatile.Read(ref _initialized) == 1,
                Environment.CurrentManagedThreadId,
                Thread.CurrentThread.GetApartmentState());
        }

        return BuildStatusSnapshot(application);
    }

    private static WpfApplication EnsureApplication()
    {
        WpfApplication? currentApplication = WpfApplication.Current;
        if (currentApplication != null)
        {
            EnsureResources(currentApplication.Resources);
            Volatile.Write(ref _initialized, 1);
            return currentApplication;
        }

        lock (SyncRoot)
        {
            currentApplication = WpfApplication.Current;
            if (currentApplication != null)
            {
                EnsureResources(currentApplication.Resources);
                Volatile.Write(ref _initialized, 1);
                AddInLog.Write("HandyControl.Application", "ReuseExisting | " + BuildStatusSnapshot(currentApplication));
                return currentApplication;
            }

            // Excel-DNA 宿主默认没有 WPF Application，这里只补齐全局资源，不额外创建隐藏窗口。
            WpfApplication application = new WpfApplication
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown
            };

            EnsureResources(application.Resources);
            Volatile.Write(ref _initialized, 1);
            AddInLog.Write("HandyControl.Application", "Created | " + BuildStatusSnapshot(application));
            return application;
        }
    }

    private static void EnsureResources(ResourceDictionary resources)
    {
        foreach (string uriText in HandyControlResourceUris)
        {
            MergeDictionary(resources, uriText);
        }

        MergeDictionary(resources, BuildLocalResourceUri(ThemeOverridesRelativePath));
        MergeDictionary(resources, BuildLocalResourceUri(DatePickerStylesRelativePath));
        EnsureDefaultAccentSnapshot(resources);
    }

    private static bool EnsureGlobalGrowlHost()
    {
        if (_globalGrowlHostWindow != null && _globalGrowlHostPanel != null)
        {
            UpdateGlobalGrowlHostBounds(_globalGrowlHostWindow);
            return true;
        }

        Rect workArea = SystemParameters.WorkArea;
        StackPanel panel = new StackPanel
        {
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Top
        };
        WpfScrollViewer scrollViewer = new WpfScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Hidden,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Background = WpfBrushes.Transparent,
            Content = panel
        };
        WpfWindow hostWindow = new WpfWindow
        {
            Width = 360,
            MaxHeight = Math.Max(200d, workArea.Height - 24d),
            Left = workArea.Right - 372d,
            Top = workArea.Top + 12d,
            WindowStyle = WindowStyle.None,
            ResizeMode = ResizeMode.NoResize,
            AllowsTransparency = true,
            Background = WpfBrushes.Transparent,
            ShowInTaskbar = false,
            ShowActivated = false,
            SizeToContent = SizeToContent.Height,
            Topmost = true,
            Content = scrollViewer
        };

        hostWindow.Closed += (_, _) =>
        {
            try
            {
                Growl.Unregister(GlobalGrowlToken);
            }
            catch
            {
            }

            _globalGrowlHostWindow = null;
            _globalGrowlHostPanel = null;
        };

        hostWindow.Show();
        Growl.Register(GlobalGrowlToken, panel);
        _globalGrowlHostWindow = hostWindow;
        _globalGrowlHostPanel = panel;
        AddInLog.Write("HandyControl.GlobalGrowlHost", "Created");
        return true;
    }

    private static void UpdateGlobalGrowlHostBounds(WpfWindow hostWindow)
    {
        Rect workArea = SystemParameters.WorkArea;
        hostWindow.MaxHeight = Math.Max(200d, workArea.Height - 24d);
        hostWindow.Left = workArea.Right - hostWindow.Width - 12d;
        hostWindow.Top = workArea.Top + 12d;
    }

    private static void MergeDictionary(ResourceDictionary resources, string uriText)
    {
        Uri uri = new Uri(uriText, UriKind.Absolute);
        MergeDictionary(resources, uri);
    }

    private static void MergeDictionary(ResourceDictionary resources, Uri uri)
    {
        bool exists = resources.MergedDictionaries.Any(item => item.Source == uri);
        if (exists)
        {
            AddInLog.Write("HandyControl.Resource", "SkippedExisting | Source=" + uri);
            return;
        }

        resources.MergedDictionaries.Add(new ResourceDictionary
        {
            Source = uri
        });
        AddInLog.Write("HandyControl.Resource", "Merged | Source=" + uri + " | TotalMerged=" + resources.MergedDictionaries.Count);
    }

    private static Uri BuildLocalResourceUri(string relativePath)
    {
        string assemblyName = typeof(HandyControlRuntime).Assembly.GetName().Name ?? "MultiTableAddin";
        string normalizedPath = relativePath.Replace("\\", "/");
        string uriText = string.Format("pack://application:,,,/{0};component/{1}", assemblyName, normalizedPath);
        return new Uri(uriText, UriKind.Absolute);
    }

    private static void EnsureDefaultAccentSnapshot(ResourceDictionary resources)
    {
        if (_defaultAccentSnapshot != null)
        {
            return;
        }

        Dictionary<string, System.Windows.Media.Color> colors = new Dictionary<string, System.Windows.Media.Color>(StringComparer.Ordinal);
        foreach ((string colorKey, _) in AccentPaletteEntries)
        {
            if (TryGetThemeColor(resources, colorKey, out System.Windows.Media.Color color))
            {
                colors[colorKey] = color;
            }
        }

        _defaultAccentSnapshot = new ThemePaletteSnapshot(colors);
    }

    private static void ApplyAccentPreset(AccentPresetOption preset)
    {
        WpfApplication application = EnsureApplication();
        ResourceDictionary resources = application.Resources;
        EnsureDefaultAccentSnapshot(resources);

        ThemePaletteSnapshot? snapshot = string.Equals(preset.Key, DefaultAccentPresetKey, StringComparison.OrdinalIgnoreCase)
            ? _defaultAccentSnapshot
            : BuildPresetSnapshot(preset);

        if (snapshot == null)
        {
            return;
        }

        foreach ((string colorKey, string brushKey) in AccentPaletteEntries)
        {
            if (!snapshot.TryGetColor(colorKey, out System.Windows.Media.Color color))
            {
                continue;
            }

            resources[colorKey] = color;
            resources[brushKey] = CreateFrozenBrush(color);
        }
    }

    private static ThemePaletteSnapshot? BuildPresetSnapshot(AccentPresetOption preset)
    {
        if (string.IsNullOrWhiteSpace(preset.PrimaryHex) ||
            string.IsNullOrWhiteSpace(preset.LightPrimaryHex) ||
            string.IsNullOrWhiteSpace(preset.DarkPrimaryHex))
        {
            return null;
        }

        Dictionary<string, System.Windows.Media.Color> colors = new Dictionary<string, System.Windows.Media.Color>(StringComparer.Ordinal)
        {
            ["PrimaryColor"] = ParseColor(preset.PrimaryHex),
            ["LightPrimaryColor"] = ParseColor(preset.LightPrimaryHex),
            ["DarkPrimaryColor"] = ParseColor(preset.DarkPrimaryHex),
            ["AccentColor"] = ParseColor(preset.PrimaryHex),
            ["DarkAccentColor"] = ParseColor(preset.DarkPrimaryHex)
        };
        return new ThemePaletteSnapshot(colors);
    }

    private static bool TryGetThemeColor(ResourceDictionary resources, string colorKey, out System.Windows.Media.Color color)
    {
        object? value = resources[colorKey];
        if (value is System.Windows.Media.Color mediaColor)
        {
            color = mediaColor;
            return true;
        }

        if (value is SolidColorBrush brush)
        {
            color = brush.Color;
            return true;
        }

        color = default;
        return false;
    }

    private static System.Windows.Media.Color ParseColor(string colorText)
    {
        return (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorText);
    }

    private static SolidColorBrush CreateFrozenBrush(System.Windows.Media.Color color)
    {
        SolidColorBrush brush = new SolidColorBrush(color);
        if (brush.CanFreeze)
        {
            brush.Freeze();
        }

        return brush;
    }

    private static string BuildStatusSnapshot(WpfApplication application)
    {
        Dispatcher dispatcher = application.Dispatcher;
        ResourceDictionary resources = application.Resources;
        int matchedResourceCount = HandyControlResourceUris.Count(uriText =>
        {
            Uri uri = new Uri(uriText, UriKind.Absolute);
            return resources.MergedDictionaries.Any(item => item.Source == uri);
        });
        bool hasThemeOverrides = resources.MergedDictionaries.Any(item => item.Source == BuildLocalResourceUri(ThemeOverridesRelativePath));

        return string.Format(
            "HandyControlState: InitializedFlag={0}; ApplicationCurrent={1}; ShutdownMode={2}; ResourceCount={3}; HandyControlResourceCount={4}; ThemeOverridesMerged={5}; DispatcherThread={6}; CurrentThread={7}; CheckAccess={8}; HasShutdownStarted={9}; HasShutdownFinished={10}; Apartment={11}",
            Volatile.Read(ref _initialized) == 1,
            ReferenceEquals(WpfApplication.Current, application),
            application.ShutdownMode,
            resources.MergedDictionaries.Count,
            matchedResourceCount,
            hasThemeOverrides,
            dispatcher.Thread.ManagedThreadId,
            Environment.CurrentManagedThreadId,
            dispatcher.CheckAccess(),
            dispatcher.HasShutdownStarted,
            dispatcher.HasShutdownFinished,
            Thread.CurrentThread.GetApartmentState());
    }

    internal sealed class AccentPresetOption
    {
        internal AccentPresetOption(string key, string displayName, string? primaryHex, string? lightPrimaryHex, string? darkPrimaryHex, string description)
        {
            Key = key;
            DisplayName = displayName;
            PrimaryHex = primaryHex;
            LightPrimaryHex = lightPrimaryHex;
            DarkPrimaryHex = darkPrimaryHex;
            Description = description;
        }

        internal string Key { get; }

        internal string DisplayName { get; }

        internal string? PrimaryHex { get; }

        internal string? LightPrimaryHex { get; }

        internal string? DarkPrimaryHex { get; }

        internal string Description { get; }
    }

    private sealed class ThemePaletteSnapshot
    {
        private readonly IReadOnlyDictionary<string, System.Windows.Media.Color> _colors;

        internal ThemePaletteSnapshot(IReadOnlyDictionary<string, System.Windows.Media.Color> colors)
        {
            _colors = colors;
        }

        internal bool TryGetColor(string colorKey, out System.Windows.Media.Color color)
        {
            return _colors.TryGetValue(colorKey, out color);
        }
    }
}
