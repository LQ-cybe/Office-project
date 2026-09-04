using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using MultiTableAddin.Core;
using UserControl = System.Windows.Controls.UserControl;

namespace MultiTableAddin.Views;

public partial class GanttView : UserControl, ITableView, IConfigAware
{
    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;
    private ViewConfigFile? _configFile;

    private const double RowHeight = 36;
    private const double HeaderHeight = 32;
    private const double GroupHeaderHeight = 28;
    private const double DayWidth = 40;
    private const double WeekWidth = 60;
    private const double MonthWidth = 90;
    private const double QuarterWidth = 120;
    private const double YearWidth = 160;

    private double _unitWidth = DayWidth;
    private double _timelineWidth;
    private DateTime _minDate;
    private DateTime _maxDate;
    private int _totalUnits;

    public GanttView()
    {
        InitializeComponent();
        Loaded += GanttView_Loaded;
    }

    private void GanttView_Loaded(object sender, RoutedEventArgs e)
    {
        // 同步左右垂直滚动
        LeftScroll.ScrollChanged += (_, _) => SyncScroll(LeftScroll, RightScroll);
        // 同步右侧滚动：左侧跟随，且冻结表头水平跟随
        RightScroll.ScrollChanged += (_, _) =>
        {
            SyncScroll(RightScroll, LeftScroll);
            HeaderCanvas.RenderTransform = new TranslateTransform(-RightScroll.HorizontalOffset, 0);
        };
    }

    private static void SyncScroll(ScrollViewer source, ScrollViewer target)
    {
        if (Math.Abs(source.VerticalOffset - target.VerticalOffset) > 0.5)
        {
            target.ScrollToVerticalOffset(source.VerticalOffset);
        }
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;

        EnsureConfig();
        UpdateDimButtons();
        BuildGantt();
    }

    public void SetConfigFile(ViewConfigFile configFile)
    {
        _configFile = configFile;
    }

    /// <summary>自动补齐并确保甘特图配置可用</summary>
    private void EnsureConfig()
    {
        _viewConfig!.GanttConfig ??= new GanttConfig();
        var cfg = _viewConfig.GanttConfig;

        var dateFields = _dataTable!.Fields.FindAll(f => f.Type == FieldType.Date).ConvertAll(f => f.Name);
        if (string.IsNullOrEmpty(cfg.StartField) && dateFields.Count > 0) cfg.StartField = dateFields[0];
        if (string.IsNullOrEmpty(cfg.EndField) && dateFields.Count > 1) cfg.EndField = dateFields[1];
        if (string.IsNullOrEmpty(cfg.LabelField) && _dataTable.Fields.Count > 0)
            cfg.LabelField = _dataTable.Fields[0].Name;
    }

    private void OnConfigClick(object sender, RoutedEventArgs e)
    {
        if (_dataTable == null || _viewConfig?.GanttConfig == null) return;

        var dialog = new GanttConfigDialog(_dataTable, _viewConfig.GanttConfig);
        dialog.Owner = Window.GetWindow(this);
        if (dialog.ShowDialog() == true)
        {
            _viewConfig.GanttConfig = dialog.Config;
            UpdateDimButtons();
            BuildGantt();
        }
    }

    /// <summary>日/周/月/季/年 切换</summary>
    private void OnDimButtonClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is TimeDimension dim && _viewConfig?.GanttConfig != null)
        {
            if (_viewConfig.GanttConfig.TimeDimension != dim)
            {
                _viewConfig.GanttConfig.TimeDimension = dim;
                UpdateDimButtons();
                BuildGantt();
            }
        }
    }

    /// <summary>高亮当前时间维度按钮</summary>
    private void UpdateDimButtons()
    {
        if (_viewConfig?.GanttConfig == null) return;
        var active = _viewConfig.GanttConfig.TimeDimension;
        foreach (Button btn in DimButtonPanel.Children)
        {
            if (btn.Tag is TimeDimension dim)
            {
                btn.Style = (Style)FindResource(dim == active ? "ButtonPrimary" : "ButtonDefault")!;
            }
        }
    }

    private void BuildGantt()
    {
        TaskTreePanel.Children.Clear();
        HeaderCanvas.Children.Clear();
        TimelineCanvas.Children.Clear();

        // 重置滚动与冻结表头偏移
        RightScroll?.ScrollToHorizontalOffset(0);
        HeaderCanvas.RenderTransform = new TranslateTransform(0, 0);

        if (_viewData == null || _viewConfig?.GanttConfig == null) return;
        var cfg = _viewConfig.GanttConfig;

        if (string.IsNullOrEmpty(cfg.StartField) || string.IsNullOrEmpty(cfg.EndField))
        {
            ShowMessage("请先在「字段设置」中指定开始日期和结束日期字段。");
            return;
        }

        var allRows = _viewData.Groups.SelectMany(g => g.Rows).ToList();
        if (allRows.Count == 0)
        {
            ShowMessage("没有可显示的数据。");
            return;
        }

        // 收集日期范围
        var dates = new List<DateTime>();
        foreach (var row in allRows)
        {
            var s = TryParseDate(row.GetValue(cfg.StartField));
            var en = TryParseDate(row.GetValue(cfg.EndField));
            if (s.HasValue) dates.Add(s.Value);
            if (en.HasValue) dates.Add(en.Value);
        }
        if (dates.Count == 0)
        {
            ShowMessage("没有有效的日期数据，请检查起止日期字段。");
            return;
        }

        _minDate = TruncateToDimension(dates.Min().AddDays(-7), cfg.TimeDimension);
        _maxDate = AddDimension(TruncateToDimension(dates.Max(), cfg.TimeDimension), cfg.TimeDimension, 2);
        _unitWidth = cfg.TimeDimension switch
        {
            TimeDimension.Year => YearWidth,
            TimeDimension.Quarter => QuarterWidth,
            TimeDimension.Month => MonthWidth,
            TimeDimension.Week => WeekWidth,
            _ => DayWidth
        };
        _totalUnits = GetUnitIndex(_maxDate, cfg.TimeDimension) - GetUnitIndex(_minDate, cfg.TimeDimension) + 1;

        double timelineWidth = Math.Max(200, _totalUnits * _unitWidth);
        _timelineWidth = timelineWidth;

        // 顶部冻结日期表头（仅一行）
        DrawTimelineHeader(HeaderCanvas, cfg.TimeDimension, timelineWidth);

        // 右侧时间轴主体
        double totalHeight = 0;
        var groups = BuildRowGroups(allRows, cfg.GroupField);
        foreach (var group in groups)
        {
            // 左侧分组头
            TaskTreePanel.Children.Add(CreateGroupHeader(group.Key, group.Rows.Count));

            // 右侧分组背景（占满宽度）
            var groupBg = new Rectangle
            {
                Width = timelineWidth,
                Height = GroupHeaderHeight,
                Fill = TryFindResource("SecondaryRegionBrush") as Brush
            };
            Canvas.SetTop(groupBg, totalHeight);
            TimelineCanvas.Children.Add(groupBg);
            totalHeight += GroupHeaderHeight;

            foreach (var row in group.Rows)
            {
                TaskTreePanel.Children.Add(CreateTaskRow(row, cfg));
                DrawTimelineRow(row, cfg, totalHeight, timelineWidth);
                totalHeight += RowHeight;
            }
        }

        TimelineCanvas.Width = timelineWidth;
        TimelineCanvas.Height = totalHeight;

        GanttInfo.Text = $"共 {allRows.Count} 项";
    }

    private void ShowMessage(string msg)
    {
        TaskTreePanel.Children.Add(new TextBlock
        {
            Text = msg,
            FontSize = 13,
            Foreground = TryFindResource("SecondaryTextBrush") as Brush,
            Margin = new Thickness(16)
        });
        GanttInfo.Text = msg;
    }

    private List<RowGroup> BuildRowGroups(List<DataRowModel> rows, string groupField)
    {
        if (string.IsNullOrEmpty(groupField))
        {
            return new List<RowGroup> { new RowGroup { Key = "全部任务", Rows = rows } };
        }

        var dict = new Dictionary<string, List<DataRowModel>>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            var key = row.GetValue(groupField)?.ToString() ?? "(空)";
            if (!dict.ContainsKey(key)) dict[key] = new List<DataRowModel>();
            dict[key].Add(row);
        }
        return dict.Select(kv => new RowGroup { Key = kv.Key, Rows = kv.Value }).ToList();
    }

    private Border CreateGroupHeader(string title, int count)
    {
        return new Border
        {
            Height = GroupHeaderHeight,
            Background = TryFindResource("SecondaryRegionBrush") as Brush,
            BorderBrush = TryFindResource("BorderBrush") as Brush,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = new TextBlock
            {
                Text = $"{title} ({count})",
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(10, 0, 8, 0),
                Foreground = TryFindResource("PrimaryTextBrush") as Brush,
                TextTrimming = TextTrimming.CharacterEllipsis
            }
        };
    }

    private Border CreateTaskRow(DataRowModel row, GanttConfig cfg)
    {
        var label = row.GetValue(cfg.LabelField)?.ToString() ?? "(空)";
        var border = new Border
        {
            Height = RowHeight,
            Background = TryFindResource("RegionBrush") as Brush,
            BorderBrush = TryFindResource("BorderBrush") as Brush,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = new TextBlock
            {
                Text = label,
                FontSize = 12,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(20, 0, 8, 0),
                Foreground = TryFindResource("PrimaryTextBrush") as Brush,
                TextTrimming = TextTrimming.CharacterEllipsis
            }
        };
        return border;
    }

    private void DrawTimelineHeader(Canvas canvas, TimeDimension dim, double timelineWidth)
    {
        canvas.Width = timelineWidth;

        // 背景
        var bg = new Rectangle
        {
            Width = timelineWidth,
            Height = HeaderHeight,
            Fill = TryFindResource("SecondaryRegionBrush") as Brush
        };
        canvas.Children.Add(bg);

        // 竖线 + 刻度文字（垂直居中）
        for (int i = 0; i < _totalUnits; i++)
        {
            var date = AddDimension(_minDate, dim, i);
            double x = i * _unitWidth;

            var line = new Line
            {
                X1 = x,
                Y1 = 0,
                X2 = x,
                Y2 = HeaderHeight,
                Stroke = TryFindResource("BorderBrush") as Brush,
                StrokeThickness = 0.5
            };
            canvas.Children.Add(line);

            string label = dim switch
            {
                TimeDimension.Year => date.ToString("yyyy年"),
                TimeDimension.Quarter => $"{date.Year}Q{(date.Month - 1) / 3 + 1}",
                TimeDimension.Month => date.ToString("yyyy-MM"),
                TimeDimension.Week => $"{date.Year}-W{ISOWeek.GetWeekOfYear(date)}",
                _ => date.ToString("MM/dd")
            };

            var tb = new TextBlock
            {
                Text = label,
                FontSize = 11,
                Foreground = TryFindResource("SecondaryTextBrush") as Brush
            };
            Canvas.SetLeft(tb, x + 4);
            Canvas.SetTop(tb, (HeaderHeight - 11) / 2.0);
            canvas.Children.Add(tb);
        }
    }

    private void DrawTimelineRow(DataRowModel row, GanttConfig cfg, double y, double timelineWidth)
    {
        // 行背景与网格线
        var rowBg = new Rectangle
        {
            Width = timelineWidth,
            Height = RowHeight,
            Fill = TryFindResource("RegionBrush") as Brush
        };
        Canvas.SetTop(rowBg, y);
        TimelineCanvas.Children.Add(rowBg);

        for (int i = 0; i < _totalUnits; i++)
        {
            double x = i * _unitWidth;
            var line = new Line
            {
                X1 = x,
                Y1 = y,
                X2 = x,
                Y2 = y + RowHeight,
                Stroke = TryFindResource("BorderBrush") as Brush,
                StrokeThickness = 0.5,
                Opacity = 0.5
            };
            TimelineCanvas.Children.Add(line);
        }

        // 甘特条
        var start = TryParseDate(row.GetValue(cfg.StartField));
        var end = TryParseDate(row.GetValue(cfg.EndField));
        if (!start.HasValue || !end.HasValue) return;
        if (end.Value < start.Value) end = start;

        double startX = GetUnitIndex(start.Value, cfg.TimeDimension) * _unitWidth;
        double endX = (GetUnitIndex(end.Value, cfg.TimeDimension) + 1) * _unitWidth;
        double barX = startX - GetUnitIndex(_minDate, cfg.TimeDimension) * _unitWidth;
        double barWidth = Math.Max(4, endX - startX);

        var bar = new Border
        {
            Width = barWidth,
            Height = RowHeight - 12,
            CornerRadius = new CornerRadius(4),
            Background = TryFindResource("PrimaryBrush") as Brush
        };
        Canvas.SetLeft(bar, barX);
        Canvas.SetTop(bar, y + 6);
        TimelineCanvas.Children.Add(bar);

        // 条内文字
        if (barWidth > 40)
        {
            var label = row.GetValue(cfg.LabelField)?.ToString() ?? "";
            var tb = new TextBlock
            {
                Text = label,
                FontSize = 10,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Padding = new Thickness(4, 0, 4, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Canvas.SetLeft(tb, barX + 4);
            Canvas.SetTop(tb, y + 10);
            TimelineCanvas.Children.Add(tb);
        }
    }

    private static double ParseProgress(object? value)
    {
        if (value == null) return 0;
        var s = value.ToString()!.Replace("%", "").Trim();
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double p))
        {
            if (p > 1 && p <= 100) p /= 100;
            return Math.Clamp(p, 0, 1);
        }
        return 0;
    }

    private static DateTime? TryParseDate(object? value)
    {
        if (value == null) return null;
        if (value is DateTime dt) return dt;
        if (DateTime.TryParse(value.ToString(), out var parsed)) return parsed;
        return null;
    }

    private static DateTime TruncateToDimension(DateTime date, TimeDimension dim) => dim switch
    {
        TimeDimension.Year => new DateTime(date.Year, 1, 1),
        TimeDimension.Quarter => new DateTime(date.Year, ((date.Month - 1) / 3) * 3 + 1, 1),
        TimeDimension.Month => new DateTime(date.Year, date.Month, 1),
        TimeDimension.Week => date.AddDays(-(int)date.DayOfWeek + 1), // 周一为周首
        _ => date.Date
    };

    private static DateTime AddDimension(DateTime date, TimeDimension dim, int count) => dim switch
    {
        TimeDimension.Year => date.AddYears(count),
        TimeDimension.Quarter => date.AddMonths(count * 3),
        TimeDimension.Month => date.AddMonths(count),
        TimeDimension.Week => date.AddDays(count * 7),
        _ => date.AddDays(count)
    };

    private int GetUnitIndex(DateTime date, TimeDimension dim) => dim switch
    {
        TimeDimension.Year => date.Year,
        TimeDimension.Quarter => date.Year * 4 + (date.Month - 1) / 3,
        TimeDimension.Month => date.Year * 12 + date.Month - 1,
        TimeDimension.Week => date.Year * 100 + ISOWeek.GetWeekOfYear(date),
        _ => (int)(date.Date - DateTime.MinValue.Date).TotalDays
    };

    private class RowGroup
    {
        public string Key { get; set; } = string.Empty;
        public List<DataRowModel> Rows { get; set; } = new();
    }
}
