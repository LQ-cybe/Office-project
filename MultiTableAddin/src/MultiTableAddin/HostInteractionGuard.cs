using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace MultiTableAddin;

internal static class HostInteractionGuard
{
    internal static bool TryBlockCommand(string stage, string caption)
    {
        if (!TryGetBusyReason(out string busyReason))
        {
            return false;
        }

        AddInLog.Write(stage + ".Blocked", "Reason=" + busyReason);
        ShowBusyAlert(caption, busyReason);
        return true;
    }

    internal static bool TryGetBusyReason(out string reason)
    {
        reason = "HostReady";

        try
        {
            object? applicationObject = ExcelDnaUtil.Application;
            if (applicationObject == null)
            {
                reason = "ApplicationNull";
                return false;
            }

            object application = applicationObject;
            object? readyValue = GetReadyValue(application);
            if (readyValue is bool ready && !ready)
            {
                reason = "Ready=False";
                return true;
            }

            object? interactiveValue = GetInteractiveValue(application);
            if (interactiveValue is bool interactive && !interactive)
            {
                reason = "Interactive=False";
                return false;
            }

            try
            {
                SetInteractiveValue(application, false);
                SetInteractiveValue(application, true);
                return false;
            }
            catch (Exception ex) when (IsRetryableOfficeBusy(ex))
            {
                reason = "InteractiveProbeBusy";
                return true;
            }
            catch
            {
                reason = "InteractiveProbeFailed";
                return true;
            }
        }
        catch (Exception ex) when (IsRetryableOfficeBusy(ex))
        {
            reason = "BusyHResult=0x" + ex.HResult.ToString("X8");
            return true;
        }
        catch
        {
            reason = "HostCheckFailed";
            return false;
        }
    }

    internal static bool IsRetryableOfficeBusy(Exception ex)
    {
        if (ex is COMException comException)
        {
            return IsRetryableOfficeBusyHResult(comException.HResult);
        }

        if (ex.HResult != 0 && IsRetryableOfficeBusyHResult(ex.HResult))
        {
            return true;
        }

        return ex.InnerException != null && IsRetryableOfficeBusy(ex.InnerException);
    }

    internal static bool IsRetryableOfficeBusyHResult(int hResult)
    {
        return hResult == unchecked((int)0x80010001)
            || hResult == unchecked((int)0x8001010A)
            || hResult == unchecked((int)0x800AC472);
    }

    private static object? GetReadyValue(object application)
    {
        return ((dynamic)application).Ready;
    }

    private static object? GetInteractiveValue(object application)
    {
        return ((dynamic)application).Interactive;
    }

    private static void SetInteractiveValue(object application, bool value)
    {
        ((dynamic)application).Interactive = value;
    }

    private static void ShowBusyAlert(string caption, string busyReason)
    {
        string message =
            caption + Environment.NewLine +
            "当前 Excel/WPS 正在忙碌，或单元格仍处于编辑状态。" + Environment.NewLine +
            "请先退出编辑、等待当前操作完成后，再重新点击。" + Environment.NewLine +
            "Reason=" + busyReason;

        try
        {
            XlCall.Excel(XlCall.xlcAlert, message);
        }
        catch (Exception ex)
        {
            AddInLog.Write("HostInteractionGuard.Alert.Error", ex.ToString());
        }
    }
}
