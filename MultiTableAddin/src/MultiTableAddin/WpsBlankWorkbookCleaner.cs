using System.Collections;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace MultiTableAddin;

internal static class WpsBlankWorkbookCleaner
{
    internal static void TryCloseBlankWorkbook(string stage)
    {
        if (!ExcelDnaUtil.IsET)
        {
            return;
        }

        try
        {
            dynamic excelApp = ExcelDnaUtil.Application;
            excelApp.Wait(DateTime.Now.AddSeconds(1));

            if ((int)excelApp.Workbooks.Count <= 1)
            {
                AddInLog.Write(stage + ".Skipped", "WorkbookCount<=1");
                return;
            }

            object? activeWorkbook = null;

            try
            {
                activeWorkbook = excelApp.ActiveWorkbook;
            }
            catch
            {
            }

            foreach (object workbookObject in (IEnumerable)excelApp.Workbooks)
            {
                dynamic workbook = workbookObject;
                bool isOtherWorkbook = !IsSameWorkbook(workbookObject, activeWorkbook);
                bool isBlankWorkbook = IsBlankWorkbook(excelApp, workbook);

                if (isOtherWorkbook && isBlankWorkbook)
                {
                    string workbookName = GetWorkbookIdentity(workbook);
                    workbook.Close(false);
                    AddInLog.Write(stage + ".Closed", workbookName);
                    return;
                }
            }

            AddInLog.Write(stage + ".Skipped", "NoBlankWorkbook");
        }
        catch (Exception ex)
        {
            AddInLog.Write(stage + ".Error", ex.ToString());
        }
    }

    private static bool IsBlankWorkbook(dynamic excelApp, dynamic workbook)
    {
        bool hasWorksheet = false;

        try
        {
            foreach (object worksheetObject in (IEnumerable)workbook.Worksheets)
            {
                hasWorksheet = true;
                if (!IsBlankWorksheet(excelApp, worksheetObject))
                {
                    return false;
                }
            }
        }
        catch
        {
            return false;
        }

        return hasWorksheet;
    }

    private static bool IsBlankWorksheet(dynamic excelApp, object worksheetObject)
    {
        object? usedRange = null;

        try
        {
            dynamic worksheet = worksheetObject;
            usedRange = worksheet.UsedRange;
            double nonEmptyCellCount = Convert.ToDouble(excelApp.WorksheetFunction.CountA(usedRange));
            return nonEmptyCellCount == 0d;
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObjectIfNeeded(usedRange);
        }
    }

    private static bool IsSameWorkbook(object? left, object? right)
    {
        if (left == null || right == null)
        {
            return false;
        }

        IntPtr leftPtr = IntPtr.Zero;
        IntPtr rightPtr = IntPtr.Zero;

        try
        {
            leftPtr = Marshal.GetIUnknownForObject(left);
            rightPtr = Marshal.GetIUnknownForObject(right);
            if (leftPtr != IntPtr.Zero && rightPtr != IntPtr.Zero)
            {
                return leftPtr == rightPtr;
            }
        }
        catch
        {
        }
        finally
        {
            if (leftPtr != IntPtr.Zero)
            {
                Marshal.Release(leftPtr);
            }

            if (rightPtr != IntPtr.Zero)
            {
                Marshal.Release(rightPtr);
            }
        }

        return string.Equals(GetWorkbookIdentity(left), GetWorkbookIdentity(right), StringComparison.OrdinalIgnoreCase);
    }

    private static string GetWorkbookIdentity(dynamic workbook)
    {
        try
        {
            string fullName = Convert.ToString(workbook.FullName) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(fullName))
            {
                return fullName;
            }
        }
        catch
        {
        }

        try
        {
            return Convert.ToString(workbook.Name) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void ReleaseComObjectIfNeeded(object? comObject)
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }
}
