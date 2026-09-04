using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace MultiTableAddin.Core;

/// <summary>
/// Excel ListObject 读写适配器
/// 隔离所有 Excel COM 操作，上层业务逻辑不直接操作 Excel 对象
/// </summary>
public class ExcelAdapter : IDisposable
{
    private dynamic? _app;
    private bool _disposed;

    private dynamic App => _app ??= ExcelDnaUtil.Application;

    private static void Release(object? com)
    {
        if (com == null) return;
        try { if (Marshal.IsComObject(com)) Marshal.ReleaseComObject(com); } catch { }
    }

    // ─────────────────────────────────────────────────────────────
    // 数据源枚举
    // ─────────────────────────────────────────────────────────────

    /// <summary>
    /// 扫描工作簿内所有工作表的超级表。
    /// 修复：旧实现只扫描 ActiveSheet，导致多个超级表时仅第一个能用。
    /// </summary>
    public List<TableSourceInfo> GetTableSources()
    {
        var result = new List<TableSourceInfo>();
        dynamic? wb = null;
        dynamic? sheets = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return result;

            sheets = wb.Worksheets;
            int sheetCount = sheets.Count;

            for (int s = 1; s <= sheetCount; s++)
            {
                dynamic? ws = null;
                dynamic? listObjects = null;
                try
                {
                    ws = sheets.Item(s);
                    if (ws == null) continue;

                    string sheetName = (string)ws.Name;
                    listObjects = ws.ListObjects;
                    if (listObjects == null) continue;

                    int loCount = listObjects.Count;
                    for (int i = 1; i <= loCount; i++)
                    {
                        dynamic? lo = null;
                        try
                        {
                            lo = listObjects.Item(i);
                            if (lo == null) continue;

                            var info = new TableSourceInfo
                            {
                                SheetName = sheetName,
                                TableName = (string)lo.Name
                            };

                            try
                            {
                                dynamic? cols = lo.ListColumns;
                                info.ColumnCount = cols == null ? 0 : (int)cols.Count;
                                Release(cols);
                            }
                            catch { }

                            try
                            {
                                dynamic? rows = lo.ListRows;
                                info.RowCount = rows == null ? 0 : (int)rows.Count;
                                Release(rows);
                            }
                            catch { }

                            result.Add(info);
                        }
                        catch (Exception exItem)
                        {
                            AddInLog.Write("ExcelAdapter.GetTableSources.Item", exItem.Message);
                        }
                        finally { Release(lo); }
                    }
                }
                catch (Exception exSheet)
                {
                    AddInLog.Write("ExcelAdapter.GetTableSources.Sheet", exSheet.Message);
                }
                finally
                {
                    Release(listObjects);
                    Release(ws);
                }
            }

            AddInLog.Write("ExcelAdapter.GetTableSources", $"Found {result.Count} tables in {sheetCount} sheets");
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.GetTableSources.Error", ex.ToString());
        }
        finally
        {
            Release(sheets);
            Release(wb);
        }
        return result;
    }

    /// <summary>获取全工作簿的所有 ListObject 名称（兼容旧调用）</summary>
    public List<string> GetListObjectNames() =>
        GetTableSources().Select(t => t.TableName).ToList();

    /// <summary>查找指定超级表所在的工作表名；找不到返回空串</summary>
    public string FindSheetOfTable(string tableName) =>
        GetTableSources().FirstOrDefault(t =>
            string.Equals(t.TableName, tableName, StringComparison.OrdinalIgnoreCase))?.SheetName
        ?? string.Empty;

    /// <summary>获取所有工作表名称</summary>
    public List<string> GetSheetNames()
    {
        var names = new List<string>();
        dynamic? wb = null;
        dynamic? sheets = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return names;

            sheets = wb.Worksheets;
            int count = sheets.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? sheet = null;
                try
                {
                    sheet = sheets.Item(i);
                    if (sheet != null) names.Add((string)sheet.Name);
                }
                finally { Release(sheet); }
            }
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.GetSheetNames.Error", ex.ToString());
        }
        finally
        {
            Release(sheets);
            Release(wb);
        }
        return names;
    }

    // ─────────────────────────────────────────────────────────────
    // 数据读取
    // ─────────────────────────────────────────────────────────────

    /// <summary>从 ListObject 读取全部数据到内存模型</summary>
    public DataTableModel ReadListObject(string sheetName, string tableName)
    {
        // 允许调用方只给表名
        if (string.IsNullOrWhiteSpace(sheetName))
            sheetName = FindSheetOfTable(tableName);

        var table = new DataTableModel
        {
            SourceFile = Path.GetFileName(GetActiveWorkbookPath()),
            SheetName = sheetName,
            TableName = tableName
        };

        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listColumns = null;
        dynamic? dataBodyRange = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return table;

            ws = wb.Worksheets[sheetName];
            if (ws == null) return table;

            lo = ws.ListObjects[tableName];
            if (lo == null) return table;

            // ── 读取列信息（含 NumberFormat，用于判定日期列）──────────
            listColumns = lo.ListColumns;
            int colCount = listColumns.Count;
            var fieldNames = new List<string>(colCount);
            var columnFormats = new List<string>(colCount);
            var columnSamples = new List<List<object?>>(colCount);

            for (int c = 1; c <= colCount; c++)
            {
                dynamic? col = null;
                try
                {
                    col = listColumns.Item(c);
                    string colName = (string)col.Name;

                    // 同名列去重，否则字典写入会互相覆盖导致「只剩几个字段」
                    string unique = colName;
                    int dup = 2;
                    while (fieldNames.Contains(unique)) unique = colName + "_" + dup++;

                    fieldNames.Add(unique);
                    columnFormats.Add(ReadColumnNumberFormat(col));
                    columnSamples.Add(new List<object?>());

                    table.Fields.Add(new FieldSchema
                    {
                        Name = unique,
                        ColumnIndex = c,
                        NumberFormat = columnFormats[c - 1]
                    });
                }
                finally { Release(col); }
            }

            // ── 批量读取数据区域 ────────────────────────────────────
            dataBodyRange = lo.DataBodyRange;
            if (dataBodyRange != null)
            {
                object?[,]? values = dataBodyRange.Value2 as object?[,];
                object?[,]? texts = null;
                try
                {
                    // 同时读取 Excel 显示文本，用于保留自定义单元格格式（如 0.0"万"、0"年"）
                    texts = dataBodyRange.Text as object?[,];
                }
                catch { }

                if (values != null)
                {
                    int rowCount = values.GetLength(0);
                    int actualColCount = values.GetLength(1);
                    int useCols = Math.Min(colCount, actualColCount);

                    for (int r = 1; r <= rowCount; r++)
                    {
                        var row = new DataRowModel { RowIndex = r };
                        for (int c = 1; c <= useCols; c++)
                        {
                            object? converted = ConvertExcelValue(values[r, c], columnFormats[c - 1]);
                            row.Values[fieldNames[c - 1]] = converted;

                            string displayText = ReadDisplayText(texts, r, c, converted, columnFormats[c - 1]);
                            row.DisplayTexts[fieldNames[c - 1]] = displayText;

                            if (columnSamples[c - 1].Count < 300)
                                columnSamples[c - 1].Add(converted);
                        }
                        table.Rows.Add(row);
                    }
                }
                else
                {
                    // 单行超级表时 Value2 不是二维数组
                    object? single = dataBodyRange.Value2;
                    object? singleText = null;
                    try { singleText = dataBodyRange.Text; } catch { }
                    if (single != null && colCount == 1)
                    {
                        object? converted = ConvertExcelValue(single, columnFormats[0]);
                        string displayText = ReadDisplayText(null, 1, 1, converted, columnFormats[0]);
                        if (singleText != null) displayText = singleText.ToString() ?? displayText;
                        table.Rows.Add(new DataRowModel
                        {
                            RowIndex = 1,
                            Values = { [fieldNames[0]] = converted },
                            DisplayTexts = { [fieldNames[0]] = displayText }
                        });
                        columnSamples[0].Add(converted);
                    }
                }
            }

            // ── 类型识别 ──────────────────────────────────────────
            var lib = FieldRuleLibrary.Current;
            for (int c = 0; c < table.Fields.Count; c++)
            {
                var detected = FieldTypeDetector.Detect(
                    table.Fields[c].Name, columnFormats[c], columnSamples[c], lib);

                table.Fields[c].Type = detected.Type;
                table.Fields[c].Options = detected.Options;
            }

            AddInLog.Write("ExcelAdapter.ReadListObject",
                $"Sheet={sheetName}, Table={tableName}, Rows={table.Rows.Count}, Cols={table.Fields.Count}");
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.ReadListObject.Error", $"{sheetName}/{tableName} :: {ex}");
        }
        finally
        {
            Release(dataBodyRange);
            Release(listColumns);
            Release(lo);
            Release(ws);
            Release(wb);
        }

        return table;
    }

    /// <summary>读取某列的数字格式；混合格式时返回空串</summary>
    private static string ReadColumnNumberFormat(dynamic col)
    {
        dynamic? range = null;
        try
        {
            range = col.DataBodyRange;
            if (range == null) return string.Empty;

            object? fmt = range.NumberFormat;
            return fmt as string ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
        finally { Release(range); }
    }

    // ─────────────────────────────────────────────────────────────
    // 数据写入
    // ─────────────────────────────────────────────────────────────

    /// <summary>更新单个单元格值</summary>
    public void UpdateCell(string sheetName, string tableName, int rowIndex, string fieldName, object? value)
    {
        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listColumns = null;
        dynamic? dataBodyRange = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return;

            ws = wb.Worksheets[sheetName];
            if (ws == null) return;

            lo = ws.ListObjects[tableName];
            if (lo == null) return;

            listColumns = lo.ListColumns;
            int colIndex = FindColumnIndex(listColumns, fieldName);
            if (colIndex < 0) return;

            dataBodyRange = lo.DataBodyRange;
            if (dataBodyRange == null) return;

            dynamic? cell = null;
            try
            {
                cell = dataBodyRange.Cells[rowIndex, colIndex];
                if (value is DateTime dt)
                {
                    cell.Value2 = dt.ToOADate();
                    string existing = cell.NumberFormat as string ?? string.Empty;
                    if (existing.Length == 0 || existing == "General" || existing == "常规")
                        cell.NumberFormat = dt.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm";
                }
                else
                {
                    cell.Value2 = ToExcelValue(value);
                }
            }
            finally { Release(cell); }
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.UpdateCell.Error", $"Row={rowIndex}, Field={fieldName}, Err={ex}");
        }
        finally
        {
            Release(dataBodyRange);
            Release(listColumns);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    /// <summary>批量更新同一行的多个字段，只做一次 COM 定位</summary>
    public void UpdateRow(string sheetName, string tableName, int rowIndex, Dictionary<string, object?> values)
    {
        if (values == null || values.Count == 0) return;

        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listColumns = null;
        dynamic? dataBodyRange = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return;

            ws = wb.Worksheets[sheetName];
            lo = ws?.ListObjects[tableName];
            if (lo == null) return;

            listColumns = lo.ListColumns;
            dataBodyRange = lo.DataBodyRange;
            if (dataBodyRange == null) return;

            var map = BuildColumnMap(listColumns);
            foreach (var kv in values)
            {
                if (!map.TryGetValue(kv.Key, out int colIndex)) continue;

                dynamic? cell = null;
                try
                {
                    cell = dataBodyRange.Cells[rowIndex, colIndex];
                    if (kv.Value is DateTime dt)
                    {
                        cell.Value2 = dt.ToOADate();
                        string existing = cell.NumberFormat as string ?? string.Empty;
                        if (existing.Length == 0 || existing == "General" || existing == "常规")
                            cell.NumberFormat = dt.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm";
                    }
                    else
                    {
                        cell.Value2 = ToExcelValue(kv.Value);
                    }
                }
                finally { Release(cell); }
            }
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.UpdateRow.Error", $"Row={rowIndex}, Err={ex}");
        }
        finally
        {
            Release(dataBodyRange);
            Release(listColumns);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    /// <summary>新增一行数据到 ListObject 末尾，返回新行号（1 基），失败返回 -1</summary>
    public int AddRow(string sheetName, string tableName, Dictionary<string, object?> values)
    {
        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listRows = null;
        dynamic? listColumns = null;
        dynamic? dataBodyRange = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return -1;

            ws = wb.Worksheets[sheetName];
            if (ws == null) return -1;

            lo = ws.ListObjects[tableName];
            if (lo == null) return -1;

            listRows = lo.ListRows;
            dynamic? newRow = null;
            try { newRow = listRows.Add(); }
            finally { Release(newRow); }

            int newRowIndex = listRows.Count;

            listColumns = lo.ListColumns;
            var map = BuildColumnMap(listColumns);

            dataBodyRange = lo.DataBodyRange;
            if (dataBodyRange != null && values != null)
            {
                foreach (var kv in values)
                {
                    if (!map.TryGetValue(kv.Key, out int colIndex)) continue;

                    dynamic? cell = null;
                    try
                    {
                        cell = dataBodyRange.Cells[newRowIndex, colIndex];
                        if (kv.Value is DateTime dt)
                        {
                            cell.Value2 = dt.ToOADate();
                            cell.NumberFormat = dt.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm";
                        }
                        else
                        {
                            cell.Value2 = ToExcelValue(kv.Value);
                        }
                    }
                    finally { Release(cell); }
                }
            }

            return newRowIndex;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.AddRow.Error", ex.ToString());
            return -1;
        }
        finally
        {
            Release(dataBodyRange);
            Release(listColumns);
            Release(listRows);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    /// <summary>在超级表末尾追加一列（若已存在同名则跳过），返回列名</summary>
    public string AddColumn(string sheetName, string tableName, string columnName)
    {
        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listColumns = null;
        dynamic? newCol = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("未找到活动工作簿");
            ws = wb.Worksheets[sheetName];
            if (ws == null) throw new InvalidOperationException("未找到工作表：" + sheetName);
            lo = ws.ListObjects[tableName];
            if (lo == null) throw new InvalidOperationException("未找到超级表：" + tableName);
            listColumns = lo.ListColumns;

            // 已存在同名列则直接返回，避免报错
            int colCount = listColumns.Count;
            for (int i = 1; i <= colCount; i++)
            {
                if (string.Equals(((string)listColumns.Item(i).Name), columnName, StringComparison.OrdinalIgnoreCase))
                    return columnName;
            }

            newCol = listColumns.Add();
            newCol.Name = columnName;
            return columnName;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.AddColumn.Error", ex.ToString());
            throw;
        }
        finally
        {
            Release(newCol);
            Release(listColumns);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    /// <summary>在指定行之前或之后插入一行</summary>
    public int InsertRow(string sheetName, string tableName, int rowIndex, bool after)
    {
        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listRows = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return -1;

            ws = wb.Worksheets[sheetName];
            if (ws == null) return -1;

            lo = ws.ListObjects[tableName];
            if (lo == null) return -1;

            listRows = lo.ListRows;
            int position = after ? rowIndex + 1 : rowIndex;
            if (position < 1) position = 1;
            if (position > listRows.Count + 1) position = listRows.Count + 1;

            dynamic? newRow = null;
            try { newRow = listRows.Add(position); }
            finally { Release(newRow); }

            return position;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.InsertRow.Error", ex.ToString());
            return -1;
        }
        finally
        {
            Release(listRows);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    /// <summary>删除一行</summary>
    public void DeleteRow(string sheetName, string tableName, int rowIndex)
    {
        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listRows = null;

        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return;

            ws = wb.Worksheets[sheetName];
            if (ws == null) return;

            lo = ws.ListObjects[tableName];
            if (lo == null) return;

            listRows = lo.ListRows;
            if (rowIndex >= 1 && rowIndex <= (int)listRows.Count)
            {
                dynamic? row = null;
                try
                {
                    row = listRows.Item(rowIndex);
                    row.Delete();
                }
                finally { Release(row); }
            }
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.DeleteRow.Error", $"Row={rowIndex}, Err={ex}");
        }
        finally
        {
            Release(listRows);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    /// <summary>选中并定位到指定数据行，方便用户在 Excel 中核对</summary>
    public void SelectRow(string sheetName, string tableName, int rowIndex)
    {
        dynamic? wb = null;
        dynamic? ws = null;
        dynamic? lo = null;
        dynamic? listRows = null;

        try
        {
            wb = App.ActiveWorkbook;
            ws = wb?.Worksheets[sheetName];
            if (ws == null) return;

            lo = ws.ListObjects[tableName];
            if (lo == null) return;

            listRows = lo.ListRows;
            if (rowIndex < 1 || rowIndex > (int)listRows.Count) return;

            dynamic? row = null;
            dynamic? range = null;
            try
            {
                ws.Activate();
                row = listRows.Item(rowIndex);
                range = row.Range;
                range.Select();
            }
            finally
            {
                Release(range);
                Release(row);
            }
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.SelectRow.Error", ex.Message);
        }
        finally
        {
            Release(listRows);
            Release(lo);
            Release(ws);
            Release(wb);
        }
    }

    // ─────────────────────────────────────────────────────────────
    // 辅助
    // ─────────────────────────────────────────────────────────────

    private static Dictionary<string, int> BuildColumnMap(dynamic listColumns)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        int colCount = listColumns.Count;
        for (int c = 1; c <= colCount; c++)
        {
            dynamic? col = null;
            try
            {
                col = listColumns.Item(c);
                string name = (string)col.Name;
                if (!map.ContainsKey(name)) map[name] = c;
            }
            catch { }
            finally { Release(col); }
        }
        return map;
    }

    private static int FindColumnIndex(dynamic listColumns, string fieldName)
    {
        int colCount = listColumns.Count;
        for (int c = 1; c <= colCount; c++)
        {
            dynamic? col = null;
            try
            {
                col = listColumns.Item(c);
                if ((string)col.Name == fieldName) return c;
            }
            catch { }
            finally { Release(col); }
        }
        return -1;
    }

    /// <summary>获取活动工作簿路径</summary>
    public string GetActiveWorkbookPath()
    {
        dynamic? wb = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb != null) return (string)wb.FullName;
        }
        catch { }
        finally { Release(wb); }
        return string.Empty;
    }

    /// <summary>获取活动工作簿名称</summary>
    public string GetActiveWorkbookName()
    {
        dynamic? wb = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb != null) return (string)wb.Name;
        }
        catch { }
        finally { Release(wb); }
        return string.Empty;
    }

    /// <summary>#428-4 配置嵌入用的隐藏工作表名（VeryHidden，普通用户不会误改）</summary>
    private const string ConfigSheetName = "_MultiTableConfig";

    /// <summary>
    /// #428-4 把配置 JSON 写入隐藏工作表（分块写入 A 列，避免单格 32767 字符上限）。
    /// 工作表不存在则创建，并设为 xlSheetVeryHidden(2) 隐藏。
    /// 返回 true 表示写入成功。
    /// </summary>
    public bool WriteConfigSheet(string json)
    {
        if (string.IsNullOrEmpty(json)) return false;
        dynamic? wb = null; dynamic? ws = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return false;
            try { ws = wb.Worksheets[ConfigSheetName]; }
            catch { ws = null; }
            if (ws == null)
            {
                ws = wb.Worksheets.Add();
                ws.Name = ConfigSheetName;
            }
            try { ws.Cells.Clear(); } catch { }
            const int chunk = 30000;
            int row = 1;
            for (int i = 0; i < json.Length; i += chunk)
            {
                int len = Math.Min(chunk, json.Length - i);
                ws.Cells[row, 1].Value = json.Substring(i, len);
                row++;
            }
            ws.Visible = 2; // xlSheetVeryHidden
            return true;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.WriteConfigSheet.Error", ex.ToString());
            return false;
        }
        finally { Release(ws); Release(wb); }
    }

    /// <summary>
    /// #428-4 从隐藏工作表读取配置 JSON；工作表不存在或为空返回 null。
    /// </summary>
    public string? ReadConfigSheet()
    {
        dynamic? wb = null; dynamic? ws = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return null;
            try { ws = wb.Worksheets[ConfigSheetName]; }
            catch { return null; }
            if (ws == null) return null;
            var sb = new System.Text.StringBuilder();
            int row = 1;
            while (row <= 10000)
            {
                var cell = ws.Cells[row, 1].Value;
                if (cell == null) break;
                string? s = cell as string;
                if (s == null) s = Convert.ToString(cell);
                if (string.IsNullOrEmpty(s)) break;
                sb.Append(s);
                row++;
            }
            return sb.Length > 0 ? sb.ToString() : null;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.ReadConfigSheet.Error", ex.ToString());
            return null;
        }
        finally { Release(ws); Release(wb); }
    }

    /// <summary>
    /// #462 当配置保存位置设为「外部文件」时删除嵌入的隐藏工作表，
    /// 以免 LoadWorkbook 隐藏表优先级加载到陈旧配置。
    /// </summary>
    public void DeleteConfigSheet()
    {
        dynamic? wb = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return;
            dynamic? ws = null;
            try { ws = wb.Worksheets[ConfigSheetName]; } catch { ws = null; }
            if (ws != null) ws.Delete();
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.DeleteConfigSheet.Error", ex.ToString());
        }
        finally { Release(wb); }
    }

    /// <summary>
    /// #466 查询嵌入配置工作表的存在性与可见性（Visible：-1=可见 / 0=普通隐藏 / 2=深度隐藏）。
    /// </summary>
    public (bool Exists, bool Visible) GetConfigSheetInfo()
    {
        dynamic? wb = null; dynamic? ws = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return (false, false);
            try { ws = wb.Worksheets[ConfigSheetName]; } catch { return (false, false); }
            if (ws == null) return (false, false);
            int vis = 2;
            try { vis = Convert.ToInt32(ws.Visible); } catch { }
            return (true, vis == -1);
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.GetConfigSheetInfo.Error", ex.ToString());
            return (false, false);
        }
        finally { Release(ws); Release(wb); }
    }

    /// <summary>
    /// #466 把配置工作表临时设为普通可见（true，并激活便于查看）或恢复深度隐藏（false）。
    /// 因为 xlSheetVeryHidden 的表不会出现在 Excel「取消隐藏」列表里，必须由程序切换。
    /// </summary>
    public bool SetConfigSheetVisible(bool visible)
    {
        dynamic? wb = null; dynamic? ws = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return false;
            try { ws = wb.Worksheets[ConfigSheetName]; } catch { return false; }
            if (ws == null) return false;
            ws.Visible = visible ? -1 : 2;   // xlSheetVisible / xlSheetVeryHidden
            if (visible)
            {
                try { App.Visible = true; } catch { }
                try { ws.Activate(); } catch { }
            }
            return true;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.SetConfigSheetVisible.Error", ex.ToString());
            return false;
        }
        finally { Release(ws); Release(wb); }
    }

    /// <summary>
    /// #466 保存当前工作簿，使嵌入的隐藏配置表真正落盘到 .xlsx（静默：DisplayAlerts=false）。
    /// 从未保存过的新工作簿（Path 为空）直接返回 false，避免弹出「另存为」对话框打扰用户。
    /// </summary>
    public bool SaveWorkbookFile()
    {
        dynamic? wb = null;
        try
        {
            wb = App.ActiveWorkbook;
            if (wb == null) return false;
            string path = string.Empty;
            try { path = Convert.ToString(wb.Path) ?? string.Empty; } catch { }
            if (string.IsNullOrEmpty(path)) return false;
            bool oldAlerts = true;
            try { oldAlerts = Convert.ToBoolean(App.DisplayAlerts); } catch { }
            try { App.DisplayAlerts = false; } catch { }
            try { wb.Save(); }
            finally { try { App.DisplayAlerts = oldAlerts; } catch { } }
            return true;
        }
        catch (Exception ex)
        {
            AddInLog.Write("ExcelAdapter.SaveWorkbookFile.Error", ex.ToString());
            return false;
        }
        finally { Release(wb); }
    }

    /// <summary>
    /// 将 Excel Value2 值转换为 C# 类型。
    /// 关键修复：Value2 会把日期返回成 OLE 序列号，需要结合列的数字格式还原为 DateTime。
    /// </summary>
    private static object? ConvertExcelValue(object? val, string numberFormat)
    {
        if (val == null) return null;

        bool formatIsDate = IsDateFormat(numberFormat);

        switch (val)
        {
            case double d:
                if (formatIsDate && ValueFormatter.TryFromOADate(d, out DateTime oaDate))
                    return oaDate;
                if (d == Math.Truncate(d) && d is > int.MinValue and < int.MaxValue)
                    return (int)d;
                return d;

            case bool b:
                return b;

            case string s:
                s = s.Trim();
                if (s.Length == 0) return null;
                // 文本形态的日期在日期列上仍然还原为 DateTime
                if (formatIsDate && ValueFormatter.TryToDateTime(s, out DateTime sd)) return sd;
                // 纯数字文本保持文本，避免编号丢失前导零
                return s;

            case DateTime dt:
                return dt;
        }

        return val;
    }

    private static bool IsDateFormat(string? fmt)
    {
        if (string.IsNullOrWhiteSpace(fmt)) return false;
        string f = fmt.Trim();
        if (f is "General" or "@" or "常规") return false;

        // 排除货币/百分比等含 m 的假阳性：要求同时出现年或日
        bool hasY = f.Contains('y') || f.Contains('年');
        bool hasD = f.Contains('d') || f.Contains('日');
        bool hasTime = f.Contains('h') || f.Contains('H');
        return hasY || (hasD && f.Contains('m')) || hasTime;
    }

    /// <summary>从批量读取的 Text 数组中获取单元格显示文本；取不到时按值与格式回退</summary>
    private static string ReadDisplayText(object?[,]? texts, int row, int col, object? converted, string numberFormat)
    {
        if (texts != null &&
            row >= 1 && row <= texts.GetLength(0) &&
            col >= 1 && col <= texts.GetLength(1))
        {
            var t = texts[row, col];
            if (t != null)
            {
                string s = t.ToString() ?? string.Empty;
                // Excel 空单元格的 Text 有时是空字符串或 "-"，这里保留非空文本
                if (!string.IsNullOrWhiteSpace(s)) return s;
            }
        }

        if (converted is DateTime dt)
            return dt.TimeOfDay == TimeSpan.Zero
                ? dt.ToString("yyyy-MM-dd")
                : dt.ToString("yyyy-MM-dd HH:mm");

        return converted?.ToString() ?? string.Empty;
    }

    /// <summary>将 C# 值转换为 Excel 可写入的值</summary>
    private static object? ToExcelValue(object? value)
    {
        if (value == null) return string.Empty;
        if (value is DateTime dt) return dt.ToOADate();
        if (value is bool b) return b;
        return value;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_app != null)
        {
            Release(_app);
            _app = null;
        }
    }
}
