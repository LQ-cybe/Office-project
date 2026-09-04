using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using MultiTableAddin.Core;
using MultiTableAddin.Html;
using MultiTableAddin.TaskPane;
using MultiTableAddin.Views;

namespace MultiTableAddin;

[ComVisible(true)]
[Guid(WpsCompatSettings.RibbonGuid)]
[ProgId(WpsCompatSettings.RibbonProgId)]
public class RibbonController : ExcelRibbon
{
    private const string RibbonOrderImageId = "RibbonOrder";
    private const string RibbonActivityImageId = "RibbonActivity";
    private const string RibbonOpenImageId = "RibbonOpen";
    private const string RibbonAboutImageId = "RibbonAbout";
    private const string RibbonXmlResourceName = "MultiTableAddin.Resources.Ribbon.xml";
    private static readonly Dictionary<string, System.Drawing.Bitmap> RibbonImageCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly object RibbonImageSyncRoot = new();

    public IRibbonUI? RibbonUI { get; private set; }

    public override string GetCustomUI(string ribbonId)
    {
        AddInLog.Write("Ribbon.GetCustomUI", ribbonId);
        return LoadRibbonXmlFromEmbeddedResource();
    }

    public override object? LoadImage(string imageId)
    {
        return TryGetRibbonImage(imageId) ?? base.LoadImage(imageId);
    }

    public object? GetButtonImage(IRibbonControl control)
    {
        string imageId = Convert.ToString(control.Tag) ?? string.Empty;
        AddInLog.Write("Ribbon.GetButtonImage", control.Id + " | Tag=" + imageId);
        return TryGetRibbonImage(imageId);
    }

    public void OnLoad(IRibbonUI ribbon)
    {
        RibbonUI = ribbon;
        AddInLog.Write("Ribbon.OnLoad");
        WpsBlankWorkbookCleaner.TryCloseBlankWorkbook("Ribbon.CloseBlankWorkbook");
    }

    public void OnShowPane(IRibbonControl control)
    {
        TaskPaneManager.Show("Ribbon.ShowPane");
    }

    public void OnShowWebPane(IRibbonControl control)
    {
        WebTaskPaneManager.Show("Ribbon.WebView2");
    }

    public void OnPaneStatus(IRibbonControl control)
    {
        TaskPaneStatusSnapshot snapshot = TaskPaneManager.GetStatusSnapshot();
        MessageBox.Show(snapshot.SummaryText, "MultiTableAddin - CTP", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void OnShowDialog(IRibbonControl control)
    {
        ExecuteRibbonAction(
            "Ribbon.ShowDialog",
            "弹窗打开失败，请先查看日志：",
            () =>
            {
                PopupDemoView popupDemoView = new PopupDemoView();
                DemoWindowService.ShowHostedWpfControl(
                    "MultiTableAddin - 示例窗口",
                    popupDemoView,
                    560,
                    860);
            });
    }

    public void OnHighlightSelection(IRibbonControl control)
    {
        ExecuteRibbonAction(
            "Ribbon.HighlightSelection",
            "高亮选区失败，请先选中单元格区域并查看日志：",
            WorkbookCommands.HighlightCurrentSelection);
    }

    public void OnRenameActiveChart(IRibbonControl control)
    {
        ExecuteRibbonAction(
            "Ribbon.RenameActiveChart",
            "图表标题更新失败，请先选中图表并查看日志：",
            () => WorkbookCommands.RenameActiveChartTitle("MultiTableAddin 示例图表"));
    }

    public void OnExportWorkbookTable(IRibbonControl control)
    {
        ExecuteRibbonAction(
            "Ribbon.ExportWorkbookTable",
            "导出结构化结果失败，请先查看日志：",
            WorkbookCommands.ExportSampleOrdersToNewSheet);
    }

    public void OnRewriteWorkbookTable(IRibbonControl control)
    {
        ExecuteRibbonAction(
            "Ribbon.RewriteWorkbookTable",
            "安全重写结构化表失败，请先查看日志：",
            () =>
            {
                WorkbookCommands.PrepareRewriteDemoTable();
                WorkbookCommands.RewritePreparedTablePreservingFormulaColumns();
            });
    }

    public void OnShowGlobalGrowl(IRibbonControl control)
    {
        string message = "MultiTableAddin 已完成全局 HandyControl 初始化，可直接从 Ribbon 调用 Growl.InfoGlobal。";
        if (HandyControlRuntime.TryInfoGlobal(message))
        {
            return;
        }

        MessageBox.Show("全局 Growl 调用失败，请先查看日志：" + AddInLog.LogFilePath, "MultiTableAddin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    private static MultiTableMainView? _mainView;
    private static FormBlankForWpf? _mainForm;

    public static MultiTableMainView? MainView => _mainView;

    public void OnOpenMultiTable(IRibbonControl control)
    {
        ExecuteRibbonAction(
            "Ribbon.OpenMultiTable",
            "打开多维表窗口失败，请先查看日志：",
            () =>
            {
                if (HtmlMainForm.Instance != null && !HtmlMainForm.Instance.IsDisposed)
                {
                    HtmlMainForm.Instance.Activate();
                    HtmlMainForm.Instance.BringToFront();
                    return;
                }

                var form = new HtmlMainForm();
                // 不以 Excel 为“拥有窗口”弹出，避免阻塞 Excel 任务栏的最小化/切换；
                // “随 Excel 最小化一起隐藏”由 HtmlMainForm 内的 WinEvent 钩子实现。
                form.Show();
            });
    }

    public void OnAbout(IRibbonControl control)
    {
        var asm = Assembly.GetExecutingAssembly();
        string dllPath = asm.Location;
        string dllModified = File.Exists(dllPath)
            ? File.GetLastWriteTime(dllPath).ToString("yyyy-MM-dd HH:mm:ss")
            : "未知";
        string message = string.Format(
            "版本：{0}\r\n编译/部署时间：{1}\r\nDLL 路径：{7}\r\n\r\nHost={2}\r\nExcelVersion={3}\r\nRibbonProgId={4}\r\nCTPProgId={5}\r\nFrameworkDescription={6}\r\nTargetFrameworkHint=.NET 8 + LatestPatch\r\nFunctions=Hello / Optional / Params / Table / RTD / Handle\r\nTaskPaneMode=Excel 2013+ 按活动窗口绑定；WPS 为单 Application 单任务窗格",
            AppVersion.DisplayText,
            dllModified,
            HostEnvironment.GetHostDisplayName(),
            ExcelDnaUtil.ExcelVersion,
            WpsCompatSettings.RibbonProgId,
            WpsCompatSettings.CtpAddInProgId,
            System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            dllPath);

        MessageBox.Show(message, "关于 MultiTableAddin", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static void ExecuteRibbonAction(string stage, string failurePrefix, Action action)
    {
        if (HostInteractionGuard.TryBlockCommand(stage, "MultiTableAddin"))
        {
            return;
        }

        try
        {
            action();
            AddInLog.Write(stage, "Success");
        }
        catch (Exception ex)
        {
            AddInLog.Write(stage + ".Error", ex.ToString());
            MessageBox.Show(failurePrefix + AddInLog.LogFilePath, "MultiTableAddin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static System.Drawing.Bitmap? TryGetRibbonImage(string imageId)
    {
        if (!string.Equals(imageId, RibbonOrderImageId, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(imageId, RibbonActivityImageId, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(imageId, RibbonOpenImageId, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(imageId, RibbonAboutImageId, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        lock (RibbonImageSyncRoot)
        {
            if (RibbonImageCache.TryGetValue(imageId, out System.Drawing.Bitmap? cachedImage))
            {
                return cachedImage;
            }

            string resourceName = "MultiTableAddin.Resources." + imageId + ".png";
            System.Drawing.Bitmap? bitmap = LoadBitmapFromEmbeddedResource(resourceName);
            if (bitmap == null)
            {
                AddInLog.Write("Ribbon.LoadImage.Missing", resourceName);
                return null;
            }

            if (HostEnvironment.IsWpsEt())
            {
                bitmap = CreateWpsTransparentSafeBitmap(bitmap);
            }

            RibbonImageCache[imageId] = bitmap;
            AddInLog.Write("Ribbon.LoadImage.Success", resourceName);
            return bitmap;
        }
    }

    private static System.Drawing.Bitmap? LoadBitmapFromEmbeddedResource(string resourceName)
    {
        Assembly assembly = typeof(RibbonController).Assembly;
        using Stream? stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            return null;
        }

        using System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
        return new System.Drawing.Bitmap(image);
    }

    private static System.Drawing.Bitmap CreateWpsTransparentSafeBitmap(System.Drawing.Image originalImage)
    {
        System.Drawing.Bitmap processedImage = new System.Drawing.Bitmap(originalImage.Width, originalImage.Height);
        using System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(processedImage);

        // 解决 WPS Ribbon 中 PNG 透明背景显示灰底的问题。
        System.Drawing.Color nearlyTransparentWhite = System.Drawing.Color.FromArgb(1, 255, 255, 255);
        graphics.Clear(nearlyTransparentWhite);
        graphics.DrawImage(originalImage, 0, 0, originalImage.Width, originalImage.Height);
        return processedImage;
    }

    private static string LoadRibbonXmlFromEmbeddedResource()
    {
        Assembly assembly = typeof(RibbonController).Assembly;
        using Stream? stream = assembly.GetManifestResourceStream(RibbonXmlResourceName);
        if (stream == null)
        {
            AddInLog.Write("Ribbon.Xml.Missing", RibbonXmlResourceName);
            throw new InvalidOperationException("未找到 Ribbon.xml 嵌入资源：" + RibbonXmlResourceName);
        }

        using StreamReader reader = new StreamReader(stream);
        string ribbonXml = reader.ReadToEnd();
        AddInLog.Write("Ribbon.Xml.Loaded", RibbonXmlResourceName);
        return ribbonXml;
    }
}
