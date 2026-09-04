using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using UserControl = System.Windows.Controls.UserControl;
using MultiTableAddin.Core;

namespace MultiTableAddin.Views;

public enum CalendarViewMode { Day, Week, Month }

public partial class CalendarView : UserControl, ITableView
{
    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;

    private DateTime _currentDate = DateTime.Now;
    private CalendarViewMode _viewMode = CalendarViewMode.Month;

    private string _dateField = "";
    private string _titleField = "";

    private static readonly string[] WeekdayNames = { "一", "二", "三", "四", "五", "六", "日" };

    public CalendarView()
    {
        InitializeComponent();
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;

        _dateField = viewConfig.CalendarConfig?.DateField ?? "";
        _titleField = viewConfig.CalendarConfig?.TitleField ?? "";

        // 尝试自动找到日期字段
        if (string.IsNullOrEmpty(_dateField))
        {
            var dateField = dataTable.Fields.Find(f => f.Type == FieldType.Date);
            if (dateField != null) _dateField = dateField.Name;
        }

        if (string.IsNullOrEmpty(_titleField) && dataTable.Fields.Count > 0)
            _titleField = dataTable.Fields[0].Name;

        // 跳到第一条记录的日期
        var allRows = viewData.Groups.SelectMany(g => g.Rows).ToList();
        if (allRows.Count > 0)
        {
            var firstDate = TryParseDate(allRows[0].GetValue(_dateField));
            if (firstDate.HasValue)
                _currentDate = firstDate.Value;
        }

        DatePicker.SelectedDate = _currentDate;
        UpdateViewButtons();
        BuildCalendar();
    }

    #region 构建日历

    private void BuildCalendar()
    {
        CalendarGrid.Children.Clear();
        CalendarGrid.RowDefinitions.Clear();
        CalendarGrid.ColumnDefinitions.Clear();

        switch (_viewMode)
        {
            case CalendarViewMode.Day:
                BuildDayView();
                break;
            case CalendarViewMode.Week:
                BuildWeekView();
                break;
            default:
                BuildMonthView();
                break;
        }
    }

    /// <summary>日视图：显示选中日期的事件列表</summary>
    private void BuildDayView()
    {
        MonthLabel.Text = _currentDate.ToString("yyyy年MM月dd日 dddd", CultureInfo.GetCultureInfo("zh-CN"));

        var records = GetRecordsForDate(_currentDate);

        if (records.Count == 0)
        {
            CalendarGrid.Children.Add(new TextBlock
            {
                Text = "当日无记录",
                FontSize = 14,
                Foreground = TryFindResource("SecondaryTextBrush") as Brush,
                Margin = new Thickness(12)
            });
            return;
        }

        var panel = new StackPanel { Margin = new Thickness(8) };
        foreach (var row in records)
        {
            panel.Children.Add(CreateRecordItem(row, true));
        }

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = panel };
        CalendarGrid.Children.Add(scroll);
    }

    /// <summary>周视图：显示一周 7 天</summary>
    private void BuildWeekView()
    {
        var startOfWeek = GetStartOfWeek(_currentDate);
        var endOfWeek = startOfWeek.AddDays(6);
        MonthLabel.Text = $"{startOfWeek:yyyy.MM.dd} - {endOfWeek:yyyy.MM.dd}";

        // 7 列
        for (int i = 0; i < 7; i++)
            CalendarGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        // 星期标题
        for (int i = 0; i < 7; i++)
        {
            var day = startOfWeek.AddDays(i);
            bool isToday = day.Date == DateTime.Today;
            var header = new Border
            {
                Background = isToday ? TryFindResource("LightPrimaryBrush") as Brush : TryFindResource("SecondaryRegionBrush") as Brush,
                BorderBrush = TryFindResource("BorderBrush") as Brush,
                BorderThickness = new Thickness(0.5),
                Padding = new Thickness(8, 6, 8, 6)
            };
            header.Child = new TextBlock
            {
                Text = $"{WeekdayNames[i]}\n{day:MM-dd}",
                FontSize = 12,
                FontWeight = isToday ? FontWeights.Bold : FontWeights.Medium,
                HorizontalAlignment = HorizontalAlignment.Center,
                TextAlignment = TextAlignment.Center,
                Foreground = isToday
                    ? TryFindResource("PrimaryBrush") as Brush
                    : TryFindResource("SecondaryTextBrush") as Brush
            };
            Grid.SetRow(header, 0);
            Grid.SetColumn(header, i);
            CalendarGrid.Children.Add(header);
        }

        // 每天记录
        for (int i = 0; i < 7; i++)
        {
            var day = startOfWeek.AddDays(i);
            var cell = CreateDayCell(day, showDayNumber: false);
            Grid.SetRow(cell, 1);
            Grid.SetColumn(cell, i);
            CalendarGrid.Children.Add(cell);
        }
    }

    /// <summary>月视图：显示整月</summary>
    private void BuildMonthView()
    {
        MonthLabel.Text = _currentDate.ToString("yyyy年MM月");

        // 7列
        for (int i = 0; i < 7; i++)
            CalendarGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // 星期标题行
        CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        for (int i = 0; i < 7; i++)
        {
            var header = new Border
            {
                Background = TryFindResource("SecondaryRegionBrush") as Brush,
                BorderBrush = TryFindResource("BorderBrush") as Brush,
                BorderThickness = new Thickness(0.5),
                Padding = new Thickness(8, 6, 8, 6)
            };
            header.Child = new TextBlock
            {
                Text = WeekdayNames[i],
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = TryFindResource("SecondaryTextBrush") as Brush
            };
            Grid.SetRow(header, 0);
            Grid.SetColumn(header, i);
            CalendarGrid.Children.Add(header);
        }

        // 计算月份天数和起始日
        var firstDayOfMonth = new DateTime(_currentDate.Year, _currentDate.Month, 1);
        int daysInMonth = DateTime.DaysInMonth(_currentDate.Year, _currentDate.Month);
        int firstDayOfWeek = (int)firstDayOfMonth.DayOfWeek;
        if (firstDayOfWeek == 0) firstDayOfWeek = 6;
        else firstDayOfWeek--;

        int totalCells = firstDayOfWeek + daysInMonth;
        int rows = (int)Math.Ceiling(totalCells / 7.0);

        for (int r = 0; r < rows; r++)
            CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        int cellIndex = 0;
        for (int day = 1; day <= daysInMonth; day++)
        {
            int row = (cellIndex + firstDayOfWeek) / 7 + 1;
            int col = (cellIndex + firstDayOfWeek) % 7;
            var date = new DateTime(_currentDate.Year, _currentDate.Month, day);

            var cell = CreateDayCell(date, showDayNumber: true);
            Grid.SetRow(cell, row);
            Grid.SetColumn(cell, col);
            CalendarGrid.Children.Add(cell);

            cellIndex++;
        }
    }

    /// <summary>创建某一天单元格</summary>
    private Border CreateDayCell(DateTime date, bool showDayNumber)
    {
        bool isToday = date.Date == DateTime.Today;
        var records = GetRecordsForDate(date);

        var cellBorder = new Border
        {
            BorderBrush = TryFindResource("BorderBrush") as Brush,
            BorderThickness = new Thickness(0.5),
            Background = isToday ? TryFindResource("LightPrimaryBrush") as Brush : TryFindResource("RegionBrush") as Brush,
            MinHeight = 80
        };

        var cellPanel = new StackPanel { Margin = new Thickness(4) };

        if (showDayNumber)
        {
            cellPanel.Children.Add(new TextBlock
            {
                Text = date.Day.ToString(),
                FontSize = 13,
                FontWeight = isToday ? FontWeights.Bold : FontWeights.Normal,
                Foreground = isToday
                    ? TryFindResource("PrimaryBrush") as Brush
                    : TryFindResource("PrimaryTextBrush") as Brush,
                Margin = new Thickness(0, 0, 0, 2)
            });
        }

        foreach (var record in records.Take(4))
        {
            cellPanel.Children.Add(CreateRecordItem(record, false));
        }

        if (records.Count > 4)
        {
            cellPanel.Children.Add(new TextBlock
            {
                Text = $"+{records.Count - 4} 更多",
                FontSize = 10,
                Foreground = TryFindResource("SecondaryTextBrush") as Brush
            });
        }

        cellBorder.Child = cellPanel;
        return cellBorder;
    }

    /// <summary>创建一条记录展示项</summary>
    private Border CreateRecordItem(DataRowModel row, bool detailed)
    {
        var titleText = row.GetValue(_titleField)?.ToString() ?? "(无标题)";
        var border = new Border
        {
            Background = TryFindResource("LightPrimaryBrush") as Brush,
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(6, 3, 6, 3),
            Margin = new Thickness(0, 0, 0, 3)
        };

        var panel = new StackPanel();
        panel.Children.Add(new TextBlock
        {
            Text = titleText,
            FontSize = 11,
            Foreground = TryFindResource("PrimaryBrush") as Brush,
            TextTrimming = TextTrimming.CharacterEllipsis
        });

        if (detailed)
        {
            foreach (var field in _dataTable?.Fields ?? new List<FieldSchema>())
            {
                if (field.Name == _titleField || field.Name == _dateField) continue;
                var val = row.GetDisplayText(field.Name, field);
                if (string.IsNullOrWhiteSpace(val)) continue;
                panel.Children.Add(new TextBlock
                {
                    Text = $"{field.Name}: {val}",
                    FontSize = 11,
                    Foreground = TryFindResource("SecondaryTextBrush") as Brush,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin = new Thickness(0, 2, 0, 0)
                });
            }
        }

        border.Child = panel;
        return border;
    }

    #endregion

    #region 数据与工具

    private List<DataRowModel> GetRecordsForDate(DateTime date)
    {
        var result = new List<DataRowModel>();
        if (_viewData == null || string.IsNullOrEmpty(_dateField)) return result;

        foreach (var row in _viewData.Groups.SelectMany(g => g.Rows))
        {
            var d = TryParseDate(row.GetValue(_dateField));
            if (d.HasValue && d.Value.Date == date.Date)
                result.Add(row);
        }
        return result;
    }

    private static DateTime? TryParseDate(object? value)
    {
        if (value == null) return null;
        if (value is DateTime dt) return dt;
        if (DateTime.TryParse(value.ToString(), out var parsed)) return parsed;
        return null;
    }

    private static DateTime GetStartOfWeek(DateTime date)
    {
        int diff = (7 + (date.DayOfWeek - DayOfWeek.Monday)) % 7;
        return date.AddDays(-1 * diff).Date;
    }

    #endregion

    #region 事件处理

    private void OnDayViewClick(object sender, RoutedEventArgs e)
    {
        _viewMode = CalendarViewMode.Day;
        UpdateViewButtons();
        BuildCalendar();
    }

    private void OnWeekViewClick(object sender, RoutedEventArgs e)
    {
        _viewMode = CalendarViewMode.Week;
        UpdateViewButtons();
        BuildCalendar();
    }

    private void OnMonthViewClick(object sender, RoutedEventArgs e)
    {
        _viewMode = CalendarViewMode.Month;
        UpdateViewButtons();
        BuildCalendar();
    }

    private void UpdateViewButtons()
    {
        void SetStyle(Button btn, bool active)
        {
            btn.Style = active
                ? (Style)FindResource("ButtonPrimary")
                : (Style)FindResource("ButtonDefault");
        }

        SetStyle(BtnDayView, _viewMode == CalendarViewMode.Day);
        SetStyle(BtnWeekView, _viewMode == CalendarViewMode.Week);
        SetStyle(BtnMonthView, _viewMode == CalendarViewMode.Month);
    }

    private void OnPrevClick(object sender, RoutedEventArgs e)
    {
        _currentDate = _viewMode switch
        {
            CalendarViewMode.Day => _currentDate.AddDays(-1),
            CalendarViewMode.Week => _currentDate.AddDays(-7),
            _ => _currentDate.AddMonths(-1)
        };
        DatePicker.SelectedDate = _currentDate;
        BuildCalendar();
    }

    private void OnNextClick(object sender, RoutedEventArgs e)
    {
        _currentDate = _viewMode switch
        {
            CalendarViewMode.Day => _currentDate.AddDays(1),
            CalendarViewMode.Week => _currentDate.AddDays(7),
            _ => _currentDate.AddMonths(1)
        };
        DatePicker.SelectedDate = _currentDate;
        BuildCalendar();
    }

    private void OnTodayClick(object sender, RoutedEventArgs e)
    {
        _currentDate = DateTime.Today;
        DatePicker.SelectedDate = _currentDate;
        BuildCalendar();
    }

    private void OnDatePickerChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (DatePicker.SelectedDate.HasValue)
        {
            _currentDate = DatePicker.SelectedDate.Value;
            BuildCalendar();
        }
    }

    #endregion
}
