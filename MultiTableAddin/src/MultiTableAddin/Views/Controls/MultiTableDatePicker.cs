using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace MultiTableAddin.Views.Controls;

/// <summary>
/// 增强型日期选择器：
/// 1) 弹出日历的“选中日期”以浅蓝色（#B3E5FC）高亮显示；
/// 2) 在日历上滚动鼠标滚轮时，每次翻动 2 个月（原生仅翻 1 个月且响应迟钝）。
///
/// 颜色通过显式设置 Calendar.CalendarDayButtonStyle 生效，而不是依赖隐式样式——
/// 隐式样式优先级过低，无法覆盖默认 Calendar 模板的选中态，这正是上一版“颜色没变”的根因。
/// </summary>
public class MultiTableDatePicker : DatePicker
{
    private const string CalendarDayButtonStyleKey = "MultiTableCalendarDayButtonStyle";

    /// <summary>每次滚轮翻动的月数（原生为 1 个月且响应迟钝，这里加速为 2 个月）。</summary>
    private const int WheelMonthsPerNotch = 2;

    public MultiTableDatePicker()
    {
        CalendarOpened += OnCalendarOpened;
        CalendarClosed += OnCalendarClosed;
    }

    private void OnCalendarOpened(object? sender, RoutedEventArgs e)
    {
        // 延迟到 Loaded 优先级，确保 Popup 与内部 Calendar 已真正加入视觉树再查找。
        // 上一版在 CalendarOpened 即时 FindChild 时 Popup 尚未就绪，导致查不到 Calendar。
        Dispatcher.BeginInvoke(
            (Action)ApplyCalendarEnhancements,
            System.Windows.Threading.DispatcherPriority.Loaded);
    }

    private void OnCalendarClosed(object? sender, RoutedEventArgs e)
    {
        if (FindCalendar() is { } calendar)
        {
            calendar.PreviewMouseWheel -= OnCalendarPreviewMouseWheel;
        }
    }

    private void ApplyCalendarEnhancements()
    {
        Calendar? calendar = FindCalendar();
        if (calendar == null)
        {
            return;
        }

        if (calendar.CalendarDayButtonStyle == null)
        {
            Style? style = ResolveStyle();
            if (style != null)
            {
                calendar.CalendarDayButtonStyle = style;
            }
        }

        // 先移除再添加，保证每次打开只绑定一次（幂等）。
        calendar.PreviewMouseWheel -= OnCalendarPreviewMouseWheel;
        calendar.PreviewMouseWheel += OnCalendarPreviewMouseWheel;
    }

    private static Style? ResolveStyle()
    {
        if (System.Windows.Application.Current?.TryFindResource(CalendarDayButtonStyleKey) is Style style)
        {
            return style;
        }

        return null;
    }

    private Calendar? FindCalendar()
    {
        ApplyTemplate();
        if (GetTemplateChild("PART_Popup") is Popup { Child: Calendar direct })
        {
            return direct;
        }

        return VisualTreeHelperEx.FindChild<Calendar>(this);
    }

    private void OnCalendarPreviewMouseWheel(object? sender, MouseWheelEventArgs e)
    {
        if (sender is not Calendar calendar)
        {
            return;
        }

        // 拦截隧道事件，阻止原生单月翻页，改为一次翻 WheelMonthsPerNotch 个月。
        e.Handled = true;
        int deltaMonths = (e.Delta > 0 ? 1 : -1) * WheelMonthsPerNotch;
        calendar.DisplayDate = calendar.DisplayDate.AddMonths(deltaMonths);
    }
}

internal static class VisualTreeHelperEx
{
    internal static T? FindChild<T>(DependencyObject? parent) where T : DependencyObject
    {
        if (parent == null)
        {
            return null;
        }

        if (parent is Popup popup && popup.Child != null)
        {
            if (popup.Child is T match)
            {
                return match;
            }

            T? result = FindChild<T>(popup.Child);
            if (result != null)
            {
                return result;
            }
        }

        int count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(parent, i);
            if (child is T directMatch)
            {
                return directMatch;
            }

            T? nested = FindChild<T>(child);
            if (nested != null)
            {
                return nested;
            }
        }

        return null;
    }
}
