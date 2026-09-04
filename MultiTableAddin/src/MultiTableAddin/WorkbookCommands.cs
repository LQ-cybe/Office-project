using System.Globalization;
using ExcelDna.Integration;

namespace MultiTableAddin;

internal static class WorkbookCommands
{
    private static readonly IReadOnlyList<ExcelColumnSchema> ExportColumnSchemas = new[]
    {
        new ExcelColumnSchema("订单号", ExcelColumnKind.Text),
        new ExcelColumnSchema("客户", ExcelColumnKind.Text),
        new ExcelColumnSchema("业务日期", ExcelColumnKind.Date),
        new ExcelColumnSchema("更新时间", ExcelColumnKind.DateTime),
        new ExcelColumnSchema("金额", ExcelColumnKind.Currency),
        new ExcelColumnSchema("完成时间", ExcelColumnKind.Time),
        new ExcelColumnSchema("状态", ExcelColumnKind.Text)
    };

    private static readonly IReadOnlyList<ExcelColumnSchema> RewriteColumnSchemas = new[]
    {
        new ExcelColumnSchema("订单号", ExcelColumnKind.Text),
        new ExcelColumnSchema("业务日期", ExcelColumnKind.Date),
        new ExcelColumnSchema("金额", ExcelColumnKind.Currency),
        new ExcelColumnSchema("完成时间", ExcelColumnKind.DateTime)
    };

    internal static void HighlightCurrentSelection()
    {
        dynamic application = ExcelDnaUtil.Application;
        object? selection = application.Selection;
        HighlightRange(selection, "Selection");
    }

    internal static void HighlightRangeByAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            throw new ArgumentException("地址不能为空。", nameof(address));
        }

        dynamic application = ExcelDnaUtil.Application;
        object? activeSheet = application.ActiveSheet;
        if (activeSheet == null)
        {
            throw new InvalidOperationException("当前没有活动工作表。");
        }

        dynamic activeSheetCom = activeSheet;
        object? range = activeSheetCom.Range[address.Trim()];
        HighlightRange(range, "Address=" + address.Trim());
    }

    internal static void RenameActiveChartTitle(string title)
    {
        object? chart = TryGetActiveChart();
        if (chart == null)
        {
            throw new InvalidOperationException("当前没有活动图表。");
        }

        RenameChartTitle(chart, title, "ActiveChart");
    }

    internal static void RenameChartTitleByName(string chartName, string title)
    {
        if (string.IsNullOrWhiteSpace(chartName))
        {
            throw new ArgumentException("图表名称不能为空。", nameof(chartName));
        }

        dynamic application = ExcelDnaUtil.Application;
        object? activeSheet = application.ActiveSheet;
        if (activeSheet == null)
        {
            throw new InvalidOperationException("当前没有活动工作表。");
        }

        dynamic activeSheetCom = activeSheet;
        object? chartObject = activeSheetCom.ChartObjects(chartName.Trim());
        object? chart = ((dynamic)chartObject).Chart;
        RenameChartTitle(chart, title, "ChartName=" + chartName.Trim());
    }

    internal static void ExportSampleOrdersToNewSheet()
    {
        dynamic application = ExcelDnaUtil.Application;
        object? activeWorkbook = application.ActiveWorkbook;
        if (activeWorkbook == null)
        {
            throw new InvalidOperationException("当前没有活动工作簿。");
        }

        using ExcelAppStateScope scope = ExcelWorkbookDataOps.BeginBulkOperation("MultiTableAddin 正在导出结构化结果...");

        object? resultWorksheet = ExcelWorkbookDataOps.CreateResultWorksheet(activeWorkbook, "自动化结果");
        ExcelWorkbookDataOps.ApplyTextColumnFormatsBeforeWrite(resultWorksheet, "A1", BuildExportRows().GetLength(0), ExportColumnSchemas);
        ExcelWorkbookDataOps.WriteTableBlock(resultWorksheet, "A1", BuildExportHeaders(), BuildExportRows());

        object? outputRange = ((dynamic)resultWorksheet).Range["A1"].CurrentRegion;
        object? listObject = ExcelWorkbookDataOps.CreateListObjectFromRange(resultWorksheet, outputRange, "tblWorkbookOrders");
        ExcelWorkbookDataOps.ApplyColumnFormats(listObject, ExportColumnSchemas);

        AddInLog.Write("WorkbookCommands.ExportSampleOrdersToNewSheet", "Sheet=自动化结果; Table=tblWorkbookOrders");
    }

    internal static void PrepareRewriteDemoTable()
    {
        dynamic application = ExcelDnaUtil.Application;
        object? activeSheet = application.ActiveSheet;
        if (activeSheet == null)
        {
            throw new InvalidOperationException("当前没有活动工作表。");
        }

        object? existingTable = ExcelWorkbookDataOps.FindListObjectByNameOnActiveSheet("tblWorkbookRewrite");
        if (existingTable != null)
        {
            ((dynamic)existingTable).Delete();
        }

        ExcelWorkbookDataOps.WriteTableBlock(activeSheet, "J1", BuildRewriteHeaders(), BuildRewriteSeedRows());
        object? outputRange = ((dynamic)activeSheet).Range["J1"].CurrentRegion;
        object? listObject = ExcelWorkbookDataOps.CreateListObjectFromRange(activeSheet, outputRange, "tblWorkbookRewrite");

        object? amountColumn = ((dynamic)listObject).ListColumns["金额"];
        object? amountRange = ((dynamic)amountColumn).DataBodyRange;
        ((dynamic)amountRange).NumberFormat = "#,##0.00";

        object? textColumn = ((dynamic)listObject).ListColumns["订单号"];
        object? textRange = ((dynamic)textColumn).DataBodyRange;
        ((dynamic)textRange).NumberFormat = "@";

        object? formulaCheckColumn = ((dynamic)listObject).ListColumns["金额校验"];
        object? formulaCheckRange = ((dynamic)formulaCheckColumn).DataBodyRange;
        ((dynamic)formulaCheckRange).FormulaR1C1 = @"=IF(RC[2]>=100,""大额"",""普通"")";

        object? formulaStatusColumn = ((dynamic)listObject).ListColumns["状态说明"];
        object? formulaStatusRange = ((dynamic)formulaStatusColumn).DataBodyRange;
        ((dynamic)formulaStatusRange).FormulaR1C1 = @"=TEXT(RC[-1],""0.00"")&""元""";

        ExcelWorkbookDataOps.ApplyColumnFormats(listObject, RewriteColumnSchemas);
        AddInLog.Write("WorkbookCommands.PrepareRewriteDemoTable", "Table=tblWorkbookRewrite");
    }

    internal static void RewritePreparedTablePreservingFormulaColumns()
    {
        object? listObject = ExcelWorkbookDataOps.FindListObjectByNameOnActiveSheet("tblWorkbookRewrite");
        if (listObject == null)
        {
            throw new InvalidOperationException("当前工作表中未找到 tblWorkbookRewrite，请先准备示例表。");
        }

        using ExcelAppStateScope scope = ExcelWorkbookDataOps.BeginBulkOperation("MultiTableAddin 正在重写 ListObject...");

        ListObjectPreserveSpec[] preserveSpecs =
        {
            new ListObjectPreserveSpec("金额校验", 3, preserveFormula: true),
            new ListObjectPreserveSpec("金额", 5, preserveNumberFormat: true),
            new ListObjectPreserveSpec("状态说明", 6, preserveFormula: true)
        };

        ExcelWorkbookDataOps.ApplyColumnFormats(listObject, RewriteColumnSchemas);
        ExcelWorkbookDataOps.RewriteListObjectPreservingColumns(listObject, BuildRewriteTargetRows(), preserveSpecs);
        ExcelWorkbookDataOps.ApplyColumnFormats(listObject, RewriteColumnSchemas);

        WorksheetTableSnapshot snapshot = ExcelWorkbookDataOps.ReadListObject(listObject);
        AddInLog.Write(
            "WorkbookCommands.RewritePreparedTablePreservingFormulaColumns",
            string.Format(
                CultureInfo.InvariantCulture,
                "Rows={0}; Cols={1}",
                snapshot.Data.GetLength(0),
                snapshot.Data.GetLength(1)));
    }

    private static void HighlightRange(object? range, string source)
    {
        if (range == null)
        {
            throw new InvalidOperationException("当前没有可操作的单元格区域。");
        }

        dynamic rangeCom = range;
        object? interior = rangeCom.Interior;
        ((dynamic)interior).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);
        string address = SafeToString(rangeCom.Address[false, false]);
        AddInLog.Write("WorkbookCommands.HighlightRange", source + "; Address=" + address);
    }

    private static void RenameChartTitle(object? chart, string title, string source)
    {
        if (chart == null)
        {
            throw new InvalidOperationException("图表对象为空。");
        }

        string safeTitle = string.IsNullOrWhiteSpace(title) ? "MultiTableAddin Smoke Chart" : title.Trim();
        dynamic chartCom = chart;
        chartCom.HasTitle = true;
        chartCom.ChartTitle.Text = safeTitle;
        AddInLog.Write("WorkbookCommands.RenameChartTitle", source + "; Title=" + safeTitle);
    }

    private static object? TryGetActiveChart()
    {
        dynamic application = ExcelDnaUtil.Application;

        try
        {
            object? activeChart = application.ActiveChart;
            if (activeChart != null)
            {
                return activeChart;
            }
        }
        catch
        {
        }

        object? selection = application.Selection;
        if (selection == null)
        {
            return null;
        }

        try
        {
            dynamic selectionCom = selection;
            object? chart = selectionCom.Chart;
            if (chart != null)
            {
                return chart;
            }
        }
        catch
        {
        }

        return null;
    }

    private static string SafeToString(object? value)
    {
        return Convert.ToString(value) ?? string.Empty;
    }

    private static object?[,] BuildExportHeaders()
    {
        return new object?[,]
        {
            { "订单号", "客户", "业务日期", "更新时间", "金额", "完成时间", "状态" }
        };
    }

    private static object?[,] BuildExportRows()
    {
        return new object?[,]
        {
            { "000000000000000123", "张三", new DateTime(2026, 7, 5), new DateTime(2026, 7, 5, 8, 30, 45), 128.5m, new DateTime(1899, 12, 30, 9, 15, 0), "已完成" },
            { "000000000000000456", "李四", new DateTime(2026, 7, 6), new DateTime(2026, 7, 6, 14, 12, 30), 86.75m, new DateTime(1899, 12, 30, 15, 45, 30), "待处理" }
        };
    }

    private static object?[,] BuildRewriteHeaders()
    {
        return new object?[,]
        {
            { "订单号", "客户", "金额校验", "业务日期", "金额", "状态说明", "完成时间" }
        };
    }

    private static object?[,] BuildRewriteSeedRows()
    {
        return new object?[,]
        {
            { "000000000000000001", "初始行", string.Empty, new DateTime(2026, 7, 1), 10m, string.Empty, new DateTime(2026, 7, 1, 9, 0, 0) }
        };
    }

    private static object?[,] BuildRewriteTargetRows()
    {
        return new object?[,]
        {
            { "000000000000000123", "更新后-张三", new DateTime(2026, 7, 7), 128.5m, new DateTime(2026, 7, 7, 9, 30, 0) },
            { "000000000000000456", "更新后-李四", new DateTime(2026, 7, 8), 86.75m, new DateTime(2026, 7, 8, 15, 45, 30) }
        };
    }

}

#if EXCELDNA_TEST_HOOKS
public static class SmokeCommands
{
    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_HIGHLIGHT_SELECTION")]
    public static void SmokeHighlightSelection()
    {
        WorkbookCommands.HighlightCurrentSelection();
    }

    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_HIGHLIGHT_ADDRESS")]
    public static void SmokeHighlightAddress(string address)
    {
        WorkbookCommands.HighlightRangeByAddress(address);
    }

    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_RENAME_ACTIVE_CHART")]
    public static void SmokeRenameActiveChart()
    {
        WorkbookCommands.RenameActiveChartTitle("MultiTableAddin Smoke Chart");
    }

    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_RENAME_CHART_BY_NAME")]
    public static void SmokeRenameChartByName(
        string chartName,
        string title = "MultiTableAddin Smoke Chart")
    {
        WorkbookCommands.RenameChartTitleByName(chartName, title);
    }

    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_EXPORT_SAMPLE_TABLE")]
    public static void SmokeExportSampleTable()
    {
        WorkbookCommands.ExportSampleOrdersToNewSheet();
    }

    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_PREPARE_REWRITE_TABLE")]
    public static void SmokePrepareRewriteTable()
    {
        WorkbookCommands.PrepareRewriteDemoTable();
    }

    [ExcelCommand(Name = "MULTITABLEADDIN_SMOKE_REWRITE_PREPARED_TABLE")]
    public static void SmokeRewritePreparedTable()
    {
        WorkbookCommands.RewritePreparedTablePreservingFormulaColumns();
    }

}
#endif
