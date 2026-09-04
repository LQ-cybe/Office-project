using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace MultiTableAddin;

internal enum ExcelColumnKind
{
    Text,
    Date,
    Time,
    DateTime,
    Number,
    Currency
}

internal sealed class ExcelColumnSchema
{
    internal ExcelColumnSchema(string columnName, ExcelColumnKind kind, string? numberFormat = null)
    {
        ColumnName = columnName;
        Kind = kind;
        NumberFormat = numberFormat;
    }

    internal string ColumnName { get; }

    internal ExcelColumnKind Kind { get; }

    internal string? NumberFormat { get; }
}

internal sealed class ListObjectPreserveSpec
{
    internal ListObjectPreserveSpec(
        string columnName,
        int targetColumnIndex,
        bool preserveFormula = false,
        bool preserveNumberFormat = false)
    {
        ColumnName = columnName;
        TargetColumnIndex = targetColumnIndex;
        PreserveFormula = preserveFormula;
        PreserveNumberFormat = preserveNumberFormat;
    }

    internal string ColumnName { get; }

    internal int TargetColumnIndex { get; }

    internal bool PreserveFormula { get; }

    internal bool PreserveNumberFormat { get; }
}

internal sealed class WorksheetTableSnapshot
{
    internal WorksheetTableSnapshot(object?[,] headers, object?[,] data)
    {
        Headers = headers;
        Data = data;
        HeaderMap = new ReadOnlyDictionary<string, int>(ExcelWorkbookDataOps.BuildHeaderMap(headers));
    }

    internal object?[,] Headers { get; }

    internal object?[,] Data { get; }

    internal IReadOnlyDictionary<string, int> HeaderMap { get; }
}

internal sealed class ExcelAppStateScope : IDisposable
{
    private const int XlCalculationManual = -4135;
    private readonly dynamic _application;
    private readonly object? _screenUpdating;
    private readonly object? _enableEvents;
    private readonly object? _displayAlerts;
    private readonly object? _calculation;
    private readonly object? _statusBar;
    private bool _disposed;

    private ExcelAppStateScope(dynamic application, string? statusMessage)
    {
        _application = application;
        _screenUpdating = TryReadProperty(() => application.ScreenUpdating);
        _enableEvents = TryReadProperty(() => application.EnableEvents);
        _displayAlerts = TryReadProperty(() => application.DisplayAlerts);
        _calculation = TryReadProperty(() => application.Calculation);
        _statusBar = TryReadProperty(() => application.StatusBar);

        TryWriteProperty(() => application.ScreenUpdating = false);
        TryWriteProperty(() => application.EnableEvents = false);
        TryWriteProperty(() => application.DisplayAlerts = false);
        TryWriteProperty(() => application.Calculation = XlCalculationManual);

        if (!string.IsNullOrWhiteSpace(statusMessage))
        {
            TryWriteProperty(() => application.StatusBar = statusMessage);
        }
    }

    internal static ExcelAppStateScope Begin(string? statusMessage = null)
    {
        dynamic application = ExcelDnaUtil.Application;
        return new ExcelAppStateScope(application, statusMessage);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        TryWriteProperty(() => _application.ScreenUpdating = _screenUpdating);
        TryWriteProperty(() => _application.EnableEvents = _enableEvents);
        TryWriteProperty(() => _application.DisplayAlerts = _displayAlerts);
        TryWriteProperty(() => _application.Calculation = _calculation);
        TryWriteProperty(() => _application.StatusBar = _statusBar ?? false);

        _disposed = true;
    }

    private static object? TryReadProperty(Func<object?> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static void TryWriteProperty(Action setter)
    {
        try
        {
            setter();
        }
        catch
        {
        }
    }
}

internal static class ExcelWorkbookDataOps
{
    private const int XlDirectionUp = -4162;
    private const int XlSrcRange = 1;
    private const int XlYes = 1;

    internal static ExcelAppStateScope BeginBulkOperation(string? statusMessage = null)
    {
        return ExcelAppStateScope.Begin(statusMessage);
    }

    internal static WorksheetTableSnapshot ReadWorksheetTable(
        object worksheet,
        int headerRow,
        object keyColumn,
        object firstColumn,
        object lastColumn)
    {
        dynamic worksheetCom = worksheet;
        object? headerRange = null;
        object? dataRange = null;

        try
        {
            int lastRow = GetLastDataRow(worksheet, keyColumn);
            headerRange = worksheetCom.Range[worksheetCom.Cells[headerRow, firstColumn], worksheetCom.Cells[headerRow, lastColumn]];
            object?[,] headers = ToManagedMatrix(((dynamic)headerRange).Value2);

            if (lastRow <= headerRow)
            {
                return new WorksheetTableSnapshot(headers, new object?[0, headers.GetLength(1)]);
            }

            dataRange = worksheetCom.Range[worksheetCom.Cells[headerRow + 1, firstColumn], worksheetCom.Cells[lastRow, lastColumn]];
            object?[,] data = ToManagedMatrix(((dynamic)dataRange).Value2);
            return new WorksheetTableSnapshot(headers, data);
        }
        finally
        {
            ReleaseComObjectIfNeeded(dataRange);
            ReleaseComObjectIfNeeded(headerRange);
        }
    }

    internal static WorksheetTableSnapshot ReadListObject(object listObject)
    {
        dynamic listObjectCom = listObject;
        object? headerRange = null;
        object? dataBodyRange = null;

        try
        {
            headerRange = listObjectCom.HeaderRowRange;
            object?[,] headers = ToManagedMatrix(((dynamic)headerRange).Value2);
            dataBodyRange = TryGetDataBodyRange(listObject);
            object?[,] data = dataBodyRange == null
                ? new object?[0, headers.GetLength(1)]
                : ToManagedMatrix(((dynamic)dataBodyRange).Value2);

            return new WorksheetTableSnapshot(headers, data);
        }
        finally
        {
            ReleaseComObjectIfNeeded(dataBodyRange);
            ReleaseComObjectIfNeeded(headerRange);
        }
    }

    internal static object?[,] NormalizeBySchema(object?[,] source, IReadOnlyDictionary<int, ExcelColumnKind> schemaByColumnIndex)
    {
        object?[,] normalized = CloneMatrix(source);
        int rowCount = normalized.GetLength(0);

        foreach (KeyValuePair<int, ExcelColumnKind> entry in schemaByColumnIndex)
        {
            int columnIndex = entry.Key;
            if (columnIndex < 1 || columnIndex > normalized.GetLength(1))
            {
                continue;
            }

            int dataColumnIndex = columnIndex - 1;
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                object? value = normalized[rowIndex, dataColumnIndex];
                if (value == null || string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.InvariantCulture)))
                {
                    continue;
                }

                normalized[rowIndex, dataColumnIndex] = NormalizeValue(value, entry.Value);
            }
        }

        return normalized;
    }

    internal static object CreateResultWorksheet(object workbook, string baseName)
    {
        dynamic workbookCom = workbook;
        object? worksheets = null;
        object? lastWorksheet = null;
        object? newWorksheet = null;

        try
        {
            worksheets = workbookCom.Worksheets;
            lastWorksheet = ((dynamic)worksheets).Item[((dynamic)worksheets).Count];
            newWorksheet = ((dynamic)worksheets).Add(Type.Missing, lastWorksheet);
            ((dynamic)newWorksheet).Name = GetUniqueWorksheetName(workbook, baseName);
            return newWorksheet;
        }
        finally
        {
            ReleaseComObjectIfNeeded(lastWorksheet);
            ReleaseComObjectIfNeeded(worksheets);
        }
    }

    internal static void WriteTableBlock(object worksheet, string topLeftAddress, object?[,] headers, object?[,] data)
    {
        dynamic worksheetCom = worksheet;
        int headerRows = headers.GetLength(0);
        int headerColumns = headers.GetLength(1);
        int dataRows = data.GetLength(0);

        if (headerRows != 1)
        {
            throw new InvalidOperationException("表头必须是单行二维数组。");
        }

        if (data.GetLength(1) != headerColumns)
        {
            throw new InvalidOperationException("数据列数必须与表头列数一致。");
        }

        object? topLeft = null;
        object? outputRange = null;
        object? headerRange = null;
        object? bodyRange = null;

        try
        {
            topLeft = worksheetCom.Range[topLeftAddress];
            outputRange = ((dynamic)topLeft).Resize[Math.Max(1, dataRows) + 1, headerColumns];
            ((dynamic)outputRange).ClearContents();

            headerRange = ((dynamic)outputRange).Rows[1];
            ((dynamic)headerRange).Value2 = ToExcelSafeArray(headers);

            if (dataRows > 0)
            {
                bodyRange = ((dynamic)outputRange).Offset[1, 0].Resize[dataRows, headerColumns];
                ((dynamic)bodyRange).Value2 = ToExcelSafeArray(data);
            }
        }
        finally
        {
            ReleaseComObjectIfNeeded(bodyRange);
            ReleaseComObjectIfNeeded(headerRange);
            ReleaseComObjectIfNeeded(outputRange);
            ReleaseComObjectIfNeeded(topLeft);
        }
    }

    internal static void ApplyTextColumnFormatsBeforeWrite(
        object worksheet,
        string topLeftAddress,
        int dataRowCount,
        IReadOnlyList<ExcelColumnSchema> schemas)
    {
        if (dataRowCount <= 0)
        {
            return;
        }

        dynamic worksheetCom = worksheet;
        object? topLeft = null;
        object? columnRange = null;

        try
        {
            topLeft = worksheetCom.Range[topLeftAddress];

            for (int columnIndex = 0; columnIndex < schemas.Count; columnIndex++)
            {
                if (schemas[columnIndex].Kind != ExcelColumnKind.Text)
                {
                    continue;
                }

                columnRange = ((dynamic)topLeft).Offset[1, columnIndex].Resize[dataRowCount, 1];
                ((dynamic)columnRange).NumberFormat = "@";
                ReleaseComObjectIfNeeded(columnRange);
                columnRange = null;
            }
        }
        finally
        {
            ReleaseComObjectIfNeeded(columnRange);
            ReleaseComObjectIfNeeded(topLeft);
        }
    }

    internal static object CreateListObjectFromRange(object worksheet, object sourceRange, string tableName)
    {
        dynamic worksheetCom = worksheet;
        object? listObjects = null;
        object? listObject = null;

        try
        {
            listObjects = worksheetCom.ListObjects;
            listObject = ((dynamic)listObjects).Add(XlSrcRange, sourceRange, Type.Missing, XlYes);
            ((dynamic)listObject).Name = tableName;
            ((dynamic)listObject).TableStyle = "TableStyleMedium2";
            return listObject;
        }
        finally
        {
            ReleaseComObjectIfNeeded(listObjects);
        }
    }

    internal static void ApplyColumnFormats(object listObject, IEnumerable<ExcelColumnSchema> schemas)
    {
        foreach (ExcelColumnSchema schema in schemas)
        {
            object? column = null;
            object? dataBodyRange = null;

            try
            {
                column = GetListColumn(listObject, schema.ColumnName);
                dataBodyRange = TryGetDataBodyRange(column);
                if (dataBodyRange == null)
                {
                    continue;
                }

                string? numberFormat = GetDefaultNumberFormat(schema);
                if (!string.IsNullOrWhiteSpace(numberFormat))
                {
                    ((dynamic)dataBodyRange).NumberFormat = numberFormat;
                }
            }
            finally
            {
                ReleaseComObjectIfNeeded(dataBodyRange);
                ReleaseComObjectIfNeeded(column);
            }
        }
    }

    internal static void RewriteListObjectPreservingColumns(
        object listObject,
        object?[,] sourceData,
        IReadOnlyList<ListObjectPreserveSpec> preserveSpecs)
    {
        int totalColumnCount = GetListObjectColumnCount(listObject);
        object?[,] writeData = BuildListObjectWriteArray(sourceData, totalColumnCount, preserveSpecs);
        IReadOnlyList<ListObjectColumnState> states = CaptureListObjectColumnStates(listObject, preserveSpecs);

        PrepareListObjectForRewrite(listObject, writeData.GetLength(0), totalColumnCount);

        object? dataBodyRange = null;

        try
        {
            dataBodyRange = TryGetDataBodyRange(listObject);
            if (dataBodyRange != null && writeData.GetLength(0) > 0)
            {
                ((dynamic)dataBodyRange).Value2 = ToExcelSafeArray(writeData);
            }
        }
        finally
        {
            ReleaseComObjectIfNeeded(dataBodyRange);
        }

        RestoreListObjectColumnStates(listObject, states);
    }

    internal static object? FindListObjectByNameOnActiveSheet(string tableName)
    {
        dynamic application = ExcelDnaUtil.Application;
        object? activeSheet = null;
        object? listObjects = null;
        object? listObject = null;

        try
        {
            activeSheet = application.ActiveSheet;
            if (activeSheet == null)
            {
                return null;
            }

            listObjects = ((dynamic)activeSheet).ListObjects;
            listObject = ((dynamic)listObjects).Item[tableName];
            return listObject;
        }
        catch
        {
            ReleaseComObjectIfNeeded(listObject);
            return null;
        }
        finally
        {
            ReleaseComObjectIfNeeded(listObjects);
            ReleaseComObjectIfNeeded(activeSheet);
        }
    }

    internal static Dictionary<string, int> BuildHeaderMap(object?[,] headerRow)
    {
        Dictionary<string, int> headerMap = new(StringComparer.OrdinalIgnoreCase);
        int columnCount = headerRow.GetLength(1);

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            string headerName = Convert.ToString(headerRow[0, columnIndex], CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(headerName) && !headerMap.ContainsKey(headerName))
            {
                headerMap[headerName] = columnIndex + 1;
            }
        }

        return headerMap;
    }

    internal static int GetLastDataRow(object worksheet, object keyColumn)
    {
        dynamic worksheetCom = worksheet;
        object? anchorCell = null;
        object? lastCell = null;

        try
        {
            anchorCell = worksheetCom.Cells[worksheetCom.Rows.Count, keyColumn];
            lastCell = ((dynamic)anchorCell).End(XlDirectionUp);
            return Convert.ToInt32(((dynamic)lastCell).Row, CultureInfo.InvariantCulture);
        }
        finally
        {
            ReleaseComObjectIfNeeded(lastCell);
            ReleaseComObjectIfNeeded(anchorCell);
        }
    }

    internal static object? CreateUnionRangeFromAddresses(object worksheet, IReadOnlyList<string> cellAddresses)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        IReadOnlyList<string> batches = SplitRangeAddressesByLength(cellAddresses);
        if (batches.Count == 0)
        {
            return null;
        }

        dynamic worksheetCom = worksheet;
        dynamic application = ExcelDnaUtil.Application;
        object? unionRange = null;

        foreach (string batch in batches)
        {
            object? batchRange = null;

            try
            {
                batchRange = worksheetCom.Range[batch];
                if (unionRange == null)
                {
                    // 首批先做一次自交集，规整重叠地址导致的异常 Areas。
                    unionRange = application.Intersect(batchRange, batchRange);
                }
                else
                {
                    unionRange = application.Union(unionRange, batchRange);
                }
            }
            finally
            {
                ReleaseComObjectIfNeeded(batchRange);
            }
        }

        return unionRange;
    }

    internal static object? CreateUnionRangeFromCellIndexes(
        object worksheet,
        IReadOnlyList<(int RowIndex, int ColIndex)> cellIndexes,
        bool checkOverload = false,
        int overloadThreshold = 5000)
    {
        IReadOnlyList<string> addresses = BuildUnionRangeAddresses(
            cellIndexes,
            checkOverload,
            overloadThreshold);

        return CreateUnionRangeFromAddresses(worksheet, addresses);
    }

    internal static IReadOnlyList<string> BuildUnionRangeAddresses(
        IReadOnlyList<(int RowIndex, int ColIndex)> cellIndexes,
        bool checkOverload = false,
        int overloadThreshold = 5000)
    {
        List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> rectangles =
            BuildCompactRectangles(cellIndexes);

        if (checkOverload && rectangles.Count > overloadThreshold)
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "待合并区域数量过大，压缩后仍有 {0} 个连续块，建议改走离线文件处理路线。",
                    rectangles.Count));
        }

        return BuildRectangleAddresses(rectangles);
    }

    private static string GetUniqueWorksheetName(object workbook, string baseName)
    {
        string candidateName = string.IsNullOrWhiteSpace(baseName) ? "结果" : baseName.Trim();
        int suffix = 1;

        while (WorksheetExists(workbook, candidateName))
        {
            suffix += 1;
            candidateName = string.Format(CultureInfo.InvariantCulture, "{0}_{1:00}", baseName, suffix);
        }

        return candidateName;
    }

    private static bool WorksheetExists(object workbook, string worksheetName)
    {
        dynamic workbookCom = workbook;
        object? worksheet = null;

        try
        {
            worksheet = workbookCom.Worksheets[worksheetName];
            return worksheet != null;
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObjectIfNeeded(worksheet);
        }
    }

    private static string? GetDefaultNumberFormat(ExcelColumnSchema schema)
    {
        if (!string.IsNullOrWhiteSpace(schema.NumberFormat))
        {
            return schema.NumberFormat;
        }

        return schema.Kind switch
        {
            ExcelColumnKind.Text => "@",
            ExcelColumnKind.Date => "yyyy-mm-dd",
            ExcelColumnKind.Time => "hh:mm:ss",
            ExcelColumnKind.DateTime => "yyyy-mm-dd hh:mm:ss",
            ExcelColumnKind.Currency => "#,##0.00",
            _ => null
        };
    }

    private static object? NormalizeValue(object value, ExcelColumnKind kind)
    {
        if (!double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
        {
            return value;
        }

        return kind switch
        {
            ExcelColumnKind.Date => DateTime.FromOADate(number).Date,
            ExcelColumnKind.Time => DateTime.FromOADate(number) - DateTime.FromOADate(Math.Floor(number)),
            ExcelColumnKind.DateTime => DateTime.FromOADate(number),
            ExcelColumnKind.Currency => Convert.ToDecimal(number, CultureInfo.InvariantCulture),
            _ => value
        };
    }

    private static IReadOnlyList<ListObjectColumnState> CaptureListObjectColumnStates(
        object listObject,
        IReadOnlyList<ListObjectPreserveSpec> preserveSpecs)
    {
        List<ListObjectColumnState> states = new();

        foreach (ListObjectPreserveSpec spec in preserveSpecs)
        {
            object? column = null;
            object? dataBodyRange = null;
            object? firstCell = null;

            try
            {
                column = GetListColumn(listObject, spec.ColumnName);
                dataBodyRange = TryGetDataBodyRange(column);

                string formulaR1C1 = string.Empty;
                string numberFormat = string.Empty;

                if (dataBodyRange != null)
                {
                    firstCell = ((dynamic)dataBodyRange).Cells[1, 1];

                    if (spec.PreserveFormula)
                    {
                        formulaR1C1 = Convert.ToString(((dynamic)firstCell).FormulaR1C1, CultureInfo.InvariantCulture) ?? string.Empty;
                    }

                    if (spec.PreserveNumberFormat)
                    {
                        numberFormat = Convert.ToString(((dynamic)firstCell).NumberFormat, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                }

                states.Add(new ListObjectColumnState(spec.ColumnName, formulaR1C1, numberFormat));
            }
            finally
            {
                ReleaseComObjectIfNeeded(firstCell);
                ReleaseComObjectIfNeeded(dataBodyRange);
                ReleaseComObjectIfNeeded(column);
            }
        }

        return states;
    }

    private static object?[,] BuildListObjectWriteArray(
        object?[,] sourceData,
        int totalColumnCount,
        IReadOnlyList<ListObjectPreserveSpec> preserveSpecs)
    {
        int rowCount = sourceData.GetLength(0);
        int sourceColumnCount = sourceData.GetLength(1);
        object?[,] result = new object?[rowCount, totalColumnCount];
        HashSet<int> reservedColumnIndexes = new(preserveSpecs.Where(spec => spec.PreserveFormula).Select(spec => spec.TargetColumnIndex));

        int sourceColumnIndex = 0;

        for (int targetColumnIndex = 1; targetColumnIndex <= totalColumnCount; targetColumnIndex++)
        {
            if (reservedColumnIndexes.Contains(targetColumnIndex))
            {
                continue;
            }

            if (sourceColumnIndex >= sourceColumnCount)
            {
                break;
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                result[rowIndex, targetColumnIndex - 1] = sourceData[rowIndex, sourceColumnIndex];
            }

            sourceColumnIndex += 1;
        }

        return result;
    }

    private static void PrepareListObjectForRewrite(object listObject, int rowCount, int totalColumnCount)
    {
        dynamic listObjectCom = listObject;
        object? dataBodyRange = null;
        object? range = null;

        try
        {
            dataBodyRange = TryGetDataBodyRange(listObject);
            if (dataBodyRange != null)
            {
                ((dynamic)dataBodyRange).ClearContents();
            }

            range = listObjectCom.Range;
            ((dynamic)listObjectCom).Resize(((dynamic)range).Resize[Math.Max(1, rowCount) + 1, totalColumnCount]);
        }
        finally
        {
            ReleaseComObjectIfNeeded(range);
            ReleaseComObjectIfNeeded(dataBodyRange);
        }
    }

    private static void RestoreListObjectColumnStates(object listObject, IReadOnlyList<ListObjectColumnState> states)
    {
        foreach (ListObjectColumnState state in states)
        {
            object? column = null;
            object? dataBodyRange = null;

            try
            {
                column = GetListColumn(listObject, state.ColumnName);
                dataBodyRange = TryGetDataBodyRange(column);
                if (dataBodyRange == null)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(state.FormulaR1C1))
                {
                    ((dynamic)dataBodyRange).FormulaR1C1 = state.FormulaR1C1;
                }

                if (!string.IsNullOrWhiteSpace(state.NumberFormat))
                {
                    ((dynamic)dataBodyRange).NumberFormat = state.NumberFormat;
                }
            }
            finally
            {
                ReleaseComObjectIfNeeded(dataBodyRange);
                ReleaseComObjectIfNeeded(column);
            }
        }
    }

    private static object? GetListColumn(object listObject, string columnName)
    {
        object? listColumns = null;
        object? column = null;

        try
        {
            listColumns = ((dynamic)listObject).ListColumns;
            column = ((dynamic)listColumns).Item[columnName];
            return column;
        }
        finally
        {
            ReleaseComObjectIfNeeded(listColumns);
        }
    }

    private static int GetListObjectColumnCount(object listObject)
    {
        object? listColumns = null;

        try
        {
            listColumns = ((dynamic)listObject).ListColumns;
            return Convert.ToInt32(((dynamic)listColumns).Count, CultureInfo.InvariantCulture);
        }
        finally
        {
            ReleaseComObjectIfNeeded(listColumns);
        }
    }

    private static object? TryGetDataBodyRange(object? listObjectOrColumn)
    {
        if (listObjectOrColumn == null)
        {
            return null;
        }

        try
        {
            return ((dynamic)listObjectOrColumn).DataBodyRange;
        }
        catch
        {
            return null;
        }
    }

    private static object?[,] CloneMatrix(object?[,] source)
    {
        int rowCount = source.GetLength(0);
        int columnCount = source.GetLength(1);
        object?[,] clone = new object?[rowCount, columnCount];

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                clone[rowIndex, columnIndex] = source[rowIndex, columnIndex];
            }
        }

        return clone;
    }

    private static List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> BuildCompactRectangles(
        IReadOnlyList<(int RowIndex, int ColIndex)> cellIndexes)
    {
        if (cellIndexes == null)
        {
            throw new ArgumentNullException(nameof(cellIndexes));
        }

        List<(int RowIndex, int ColIndex)> normalizedCells = cellIndexes
            .Where(cell => cell.RowIndex > 0 && cell.ColIndex > 0)
            .Distinct()
            .OrderBy(cell => cell.RowIndex)
            .ThenBy(cell => cell.ColIndex)
            .ToList();

        if (normalizedCells.Count == 0)
        {
            return new List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)>();
        }

        return ChooseBetterRectanglePlan(
            BuildRectanglesByRowFirst(normalizedCells),
            BuildRectanglesByColumnFirst(normalizedCells));
    }

    private static List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> BuildRectanglesByRowFirst(
        IReadOnlyList<(int RowIndex, int ColIndex)> normalizedCells)
    {
        List<(int RowIndex, int FirstColIndex, int LastColIndex)> horizontalRuns = new();

        foreach (IGrouping<int, (int RowIndex, int ColIndex)> rowGroup in normalizedCells.GroupBy(cell => cell.RowIndex).OrderBy(group => group.Key))
        {
            foreach ((int firstIndex, int lastIndex) run in BuildContinuousRuns(rowGroup.Select(cell => cell.ColIndex)))
            {
                horizontalRuns.Add((rowGroup.Key, run.firstIndex, run.lastIndex));
            }
        }

        List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> rectangles = new();
        foreach (IGrouping<(int FirstColIndex, int LastColIndex), (int RowIndex, int FirstColIndex, int LastColIndex)> group in
                 horizontalRuns.GroupBy(run => (run.FirstColIndex, run.LastColIndex)))
        {
            foreach ((int firstIndex, int lastIndex) run in BuildContinuousRuns(group.Select(item => item.RowIndex)))
            {
                rectangles.Add((run.firstIndex, run.lastIndex, group.Key.FirstColIndex, group.Key.LastColIndex));
            }
        }

        return rectangles;
    }

    private static List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> BuildRectanglesByColumnFirst(
        IReadOnlyList<(int RowIndex, int ColIndex)> normalizedCells)
    {
        List<(int ColIndex, int FirstRowIndex, int LastRowIndex)> verticalRuns = new();

        foreach (IGrouping<int, (int RowIndex, int ColIndex)> columnGroup in normalizedCells.GroupBy(cell => cell.ColIndex).OrderBy(group => group.Key))
        {
            foreach ((int firstIndex, int lastIndex) run in BuildContinuousRuns(columnGroup.Select(cell => cell.RowIndex)))
            {
                verticalRuns.Add((columnGroup.Key, run.firstIndex, run.lastIndex));
            }
        }

        List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> rectangles = new();
        foreach (IGrouping<(int FirstRowIndex, int LastRowIndex), (int ColIndex, int FirstRowIndex, int LastRowIndex)> group in
                 verticalRuns.GroupBy(run => (run.FirstRowIndex, run.LastRowIndex)))
        {
            foreach ((int firstIndex, int lastIndex) run in BuildContinuousRuns(group.Select(item => item.ColIndex)))
            {
                rectangles.Add((group.Key.FirstRowIndex, group.Key.LastRowIndex, run.firstIndex, run.lastIndex));
            }
        }

        return rectangles;
    }

    private static List<(int firstIndex, int lastIndex)> BuildContinuousRuns(IEnumerable<int> indexes)
    {
        List<int> ordered = indexes.Distinct().OrderBy(index => index).ToList();
        List<(int firstIndex, int lastIndex)> runs = new();
        if (ordered.Count == 0)
        {
            return runs;
        }

        int start = ordered[0];
        int previous = ordered[0];

        for (int i = 1; i < ordered.Count; i++)
        {
            int current = ordered[i];
            if (current == previous + 1)
            {
                previous = current;
                continue;
            }

            runs.Add((start, previous));
            start = current;
            previous = current;
        }

        runs.Add((start, previous));
        return runs;
    }

    private static List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> ChooseBetterRectanglePlan(
        List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> rowFirstRectangles,
        List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> columnFirstRectangles)
    {
        (int RectangleCount, int BatchCount, int TotalAddressLength) rowFirstScore = EvaluateRectanglePlan(rowFirstRectangles);
        (int RectangleCount, int BatchCount, int TotalAddressLength) columnFirstScore = EvaluateRectanglePlan(columnFirstRectangles);

        int rectangleComparison = rowFirstScore.RectangleCount.CompareTo(columnFirstScore.RectangleCount);
        if (rectangleComparison < 0)
        {
            return rowFirstRectangles;
        }

        if (rectangleComparison > 0)
        {
            return columnFirstRectangles;
        }

        int batchComparison = rowFirstScore.BatchCount.CompareTo(columnFirstScore.BatchCount);
        if (batchComparison < 0)
        {
            return rowFirstRectangles;
        }

        if (batchComparison > 0)
        {
            return columnFirstRectangles;
        }

        int addressLengthComparison = rowFirstScore.TotalAddressLength.CompareTo(columnFirstScore.TotalAddressLength);
        return addressLengthComparison <= 0 ? rowFirstRectangles : columnFirstRectangles;
    }

    private static (int RectangleCount, int BatchCount, int TotalAddressLength) EvaluateRectanglePlan(
        List<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> rectangles)
    {
        IReadOnlyList<string> addresses = BuildRectangleAddresses(rectangles);
        int totalAddressLength = addresses.Sum(address => address.Length);
        int batchCount = SplitRangeAddressesByLength(addresses).Count;
        return (rectangles.Count, batchCount, totalAddressLength);
    }

    private static IReadOnlyList<string> BuildRectangleAddresses(
        IEnumerable<(int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex)> rectangles)
    {
        return rectangles
            .Select(FormatRectangleAddress)
            .ToList();
    }

    private static string FormatRectangleAddress((int FirstRowIndex, int LastRowIndex, int FirstColIndex, int LastColIndex) rectangle)
    {
        string firstAddress = string.Format(
            CultureInfo.InvariantCulture,
            "{0}{1}",
            ToExcelColumnName(rectangle.FirstColIndex),
            rectangle.FirstRowIndex);

        if (rectangle.FirstRowIndex == rectangle.LastRowIndex && rectangle.FirstColIndex == rectangle.LastColIndex)
        {
            return firstAddress;
        }

        return string.Format(
            CultureInfo.InvariantCulture,
            "{0}:{1}{2}",
            firstAddress,
            ToExcelColumnName(rectangle.LastColIndex),
            rectangle.LastRowIndex);
    }

    private static IReadOnlyList<string> SplitRangeAddressesByLength(IReadOnlyList<string> cellAddresses, int maxLength = 255)
    {
        if (cellAddresses == null)
        {
            throw new ArgumentNullException(nameof(cellAddresses));
        }

        List<string> result = new();
        List<string> currentBatch = new();
        int currentLength = 0;

        foreach (string address in cellAddresses.Where(address => !string.IsNullOrWhiteSpace(address)))
        {
            int nextLength = currentLength == 0 ? address.Length : currentLength + 1 + address.Length;
            if (currentBatch.Count > 0 && nextLength > maxLength)
            {
                result.Add(string.Join(",", currentBatch));
                currentBatch.Clear();
                currentLength = 0;
            }

            currentBatch.Add(address);
            currentLength = currentLength == 0 ? address.Length : currentLength + 1 + address.Length;
        }

        if (currentBatch.Count > 0)
        {
            result.Add(string.Join(",", currentBatch));
        }

        return result;
    }

    private static string ToExcelColumnName(int columnIndex)
    {
        if (columnIndex <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        }

        int dividend = columnIndex;
        List<char> chars = new();

        while (dividend > 0)
        {
            dividend--;
            chars.Add((char)('A' + (dividend % 26)));
            dividend /= 26;
        }

        chars.Reverse();
        return new string(chars.ToArray());
    }

    private static object?[,] ToManagedMatrix(object? excelValue)
    {
        if (excelValue is Array excelArray && excelArray.Rank == 2)
        {
            int rowCount = excelArray.GetLength(0);
            int columnCount = excelArray.GetLength(1);
            int rowLowerBound = excelArray.GetLowerBound(0);
            int columnLowerBound = excelArray.GetLowerBound(1);
            object?[,] managed = new object?[rowCount, columnCount];

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    managed[rowIndex, columnIndex] = excelArray.GetValue(rowIndex + rowLowerBound, columnIndex + columnLowerBound);
                }
            }

            return managed;
        }

        return new object?[1, 1] { { excelValue } };
    }

    private static Array ToExcelSafeArray(object?[,] managed)
    {
        int rowCount = managed.GetLength(0);
        int columnCount = managed.GetLength(1);
        Array safeArray = Array.CreateInstance(typeof(object), new[] { rowCount, columnCount }, new[] { 1, 1 });

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                safeArray.SetValue(managed[rowIndex, columnIndex], rowIndex + 1, columnIndex + 1);
            }
        }

        return safeArray;
    }

    internal static void ReleaseComObjectIfNeeded(object? comObject)
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }

    private sealed class ListObjectColumnState
    {
        internal ListObjectColumnState(string columnName, string formulaR1C1, string numberFormat)
        {
            ColumnName = columnName;
            FormulaR1C1 = formulaR1C1;
            NumberFormat = numberFormat;
        }

        internal string ColumnName { get; }

        internal string FormulaR1C1 { get; }

        internal string NumberFormat { get; }
    }
}
