using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using UserControl = System.Windows.Controls.UserControl;
using System.Windows.Media;
using System.Windows.Media.Effects;
using MultiTableAddin.Core;

namespace MultiTableAddin.Views;

public partial class KanbanView : UserControl, ITableView, IConfigAware
{
    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;
    private ViewConfigFile? _configFile;
    private DataRowModel? _draggingRow;
    private Border? _selectedCard;

    public KanbanView()
    {
        InitializeComponent();
    }

    public void SetConfigFile(ViewConfigFile configFile)
    {
        _configFile = configFile;
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;
        BuildToolbar();
        BuildColumns();
    }

    /// <summary>构建顶部工具栏选项</summary>
    private void BuildToolbar()
    {
        if (_dataTable == null || _viewConfig == null) return;

        var fieldNames = _dataTable.Fields.Select(f => f.Name).ToList();

        // 分组字段：任意字段均可分组，空表示不分组
        GroupByCombo.SelectionChanged -= OnGroupByChanged;
        GroupByCombo.ItemsSource = new List<string> { "(不分组)" }.Concat(fieldNames).ToList();
        GroupByCombo.SelectedItem = string.IsNullOrEmpty(_viewConfig.GroupBy)
            ? "(不分组)"
            : _viewConfig.GroupBy;
        GroupByCombo.SelectionChanged += OnGroupByChanged;

        // 排序字段
        SortFieldCombo.ItemsSource = new List<string> { "(不排序)" }.Concat(fieldNames).ToList();
        if (_viewConfig.Sort.Count > 0)
            SortFieldCombo.SelectedItem = _viewConfig.Sort[0].Field;
        else
            SortFieldCombo.SelectedItem = "(不排序)";

        SortOrderCombo.ItemsSource = new List<string> { "升序", "降序" };
        SortOrderCombo.SelectedItem = _viewConfig.Sort.FirstOrDefault()?.Order?.Equals("desc", StringComparison.OrdinalIgnoreCase) == true
            ? "降序"
            : "升序";
    }

    private void BuildColumns()
    {
        ColumnsPanel.Children.Clear();
        if (_viewData == null || _viewConfig == null || _dataTable == null) return;

        foreach (var group in _viewData.Groups)
        {
            var columnBorder = new Border
            {
                Style = (Style)FindResource("KanbanColumn"),
                Tag = group.Key
            };

            // 列整体：顶部固定列头 + 下方可垂直滚动的卡片区
            var columnPanel = new DockPanel { Margin = new Thickness(8) };

            // 列头（固定，不随卡片滚动）
            var headerPanel = new DockPanel { Margin = new Thickness(4, 4, 4, 8) };
            var titleText = new TextBlock
            {
                Text = $"{group.Key} ({group.Rows.Count})",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = TryFindResource("PrimaryTextBrush") as Brush
            };
            DockPanel.SetDock(titleText, Dock.Left);
            headerPanel.Children.Add(titleText);
            DockPanel.SetDock(headerPanel, Dock.Top);
            columnPanel.Children.Add(headerPanel);

            // 卡片区：独立垂直滚动，数据多时可滚动查看
            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };
            var cardsPanel = new StackPanel();
            foreach (var row in group.Rows)
            {
                var card = CreateCard(row);
                cardsPanel.Children.Add(card);
            }

            // 拖放占位区
            var dropZone = new Border
            {
                Height = 60,
                Background = Brushes.Transparent,
                Tag = group.Key,
                Margin = new Thickness(0, 0, 0, 8)
            };
            cardsPanel.Children.Add(dropZone);

            scroll.Content = cardsPanel;
            columnPanel.Children.Add(scroll);

            columnBorder.Child = columnPanel;

            // 拖放事件
            columnBorder.AllowDrop = true;
            columnBorder.DragEnter += Column_DragEnter;
            columnBorder.DragLeave += Column_DragLeave;
            columnBorder.Drop += Column_Drop;

            ColumnsPanel.Children.Add(columnBorder);
        }
    }

    private Border CreateCard(DataRowModel row)
    {
        var cardBorder = new Border
        {
            Style = (Style)FindResource("KanbanCard"),
            Tag = row
        };

        var cardPanel = new StackPanel();

        // 标题
        string titleField = _viewConfig?.CardMeta?.Title ?? "";
        if (string.IsNullOrEmpty(titleField) && _dataTable?.Fields.Count > 0)
            titleField = _dataTable.Fields[0].Name;

        var titleText = new TextBlock
        {
            Text = row.GetValue(titleField)?.ToString() ?? "(空)",
            FontSize = 14,
            FontWeight = FontWeights.Medium,
            Foreground = TryFindResource("PrimaryTextBrush") as Brush,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(0, 0, 0, 4)
        };
        cardPanel.Children.Add(titleText);

        // 描述字段
        if (_viewConfig?.CardMeta?.Description != null)
        {
            foreach (var descField in _viewConfig.CardMeta.Description)
            {
                var val = row.GetValue(descField);
                if (val == null) continue;
                var descText = new TextBlock
                {
                    Text = $"{descField}: {val}",
                    FontSize = 12,
                    Foreground = TryFindResource("SecondaryTextBrush") as Brush,
                    Margin = new Thickness(0, 2, 0, 0),
                    TextTrimming = TextTrimming.CharacterEllipsis
                };
                cardPanel.Children.Add(descText);
            }
        }

        cardBorder.Child = cardPanel;

        // 拖拽支持
        cardBorder.MouseMove += Card_MouseMove;
        cardBorder.MouseLeftButtonDown += Card_MouseLeftButtonDown;
        cardBorder.MouseLeftButtonUp += Card_MouseLeftButtonUp;

        return cardBorder;
    }

    private void Card_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Border card) return;

        // 恢复之前选中的卡片
        if (_selectedCard != null && _selectedCard != card)
        {
            _selectedCard.BorderBrush = TryFindResource("BorderBrush") as Brush;
            _selectedCard.BorderThickness = new Thickness(1);
            _selectedCard.Background = TryFindResource("RegionBrush") as Brush;
        }

        // 高亮当前卡片
        _selectedCard = card;
        _selectedCard.BorderBrush = TryFindResource("PrimaryBrush") as Brush;
        _selectedCard.BorderThickness = new Thickness(2);
        _selectedCard.Background = TryFindResource("LightPrimaryBrush") as Brush ?? TryFindResource("RegionBrush") as Brush;

        e.Handled = true;
    }

    private void Card_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border border && border.Tag is DataRowModel row)
        {
            _draggingRow = row;
        }
    }

    private void Card_MouseMove(object sender, MouseEventArgs e)
    {
        if (_draggingRow == null) return;
        if (e.LeftButton == MouseButtonState.Pressed)
        {
            if (sender is Border border)
            {
                DragDrop.DoDragDrop(border, _draggingRow, DragDropEffects.Move);
            }
        }
        else
        {
            _draggingRow = null;
        }
    }

    private void Column_DragEnter(object sender, DragEventArgs e)
    {
        if (sender is Border border)
        {
            border.Background = TryFindResource("LightPrimaryBrush") as Brush ?? new SolidColorBrush(Color.FromRgb(230, 240, 255));
        }
        e.Effects = DragDropEffects.Move;
    }

    private void Column_DragLeave(object sender, DragEventArgs e)
    {
        if (sender is Border border)
        {
            border.Background = TryFindResource("SecondaryRegionBrush") as Brush;
        }
    }

    private void Column_Drop(object sender, DragEventArgs e)
    {
        if (sender is Border border && _draggingRow != null && _dataTable != null && _excelAdapter != null)
        {
            string newGroupKey = border.Tag?.ToString() ?? "";
            string groupField = _viewConfig?.GroupBy ?? "";

            if (!string.IsNullOrEmpty(groupField))
            {
                // 更新内存数据
                _draggingRow.SetValue(groupField, newGroupKey);
                // 回写到 Excel
                _excelAdapter.UpdateCell(_dataTable.SheetName, _dataTable.TableName, _draggingRow.RowIndex, groupField, newGroupKey);
                _dataTable.IsDirty = true;

                AddInLog.Write("KanbanView.DragDrop", $"Row={_draggingRow.RowIndex}, Field={groupField}, Value={newGroupKey}");
            }
        }

        // 恢复列背景
        if (sender is Border b)
        {
            b.Background = TryFindResource("SecondaryRegionBrush") as Brush;
        }

        _draggingRow = null;

        // 重新构建列
        if (_dataTable != null && _viewConfig != null)
        {
            var newViewData = new ViewEngine().Apply(_dataTable, _viewConfig);
            _viewData = newViewData;
            BuildColumns();
        }
    }

    /// <summary>分组字段切换</summary>
    private void OnGroupByChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_viewConfig == null || _dataTable == null) return;

        var selected = GroupByCombo.SelectedItem as string;
        _viewConfig.GroupBy = selected == "(不分组)" ? string.Empty : selected ?? string.Empty;

        RefreshViewData();
        SaveConfig();
    }

    /// <summary>字段显隐：选择卡片描述字段</summary>
    private void OnFieldVisibilityClick(object sender, RoutedEventArgs e)
    {
        if (_viewConfig == null || _dataTable == null) return;

        var current = _viewConfig.CardMeta?.Description ?? new List<string>();
        var picker = new FieldPickerDialog(_dataTable.Fields.Select(f => f.Name).ToList(), current, "选择卡片显示字段")
        {
            Owner = Window.GetWindow(this)
        };

        if (picker.ShowDialog() == true)
        {
            _viewConfig.CardMeta ??= new CardMeta();
            _viewConfig.CardMeta.Description = new List<string>(picker.SelectedFields);
            BuildColumns();
            SaveConfig();
        }
    }

    /// <summary>应用排序</summary>
    private void OnApplySortClick(object sender, RoutedEventArgs e)
    {
        if (_viewConfig == null || _dataTable == null) return;

        var field = SortFieldCombo.SelectedItem as string;
        var orderText = SortOrderCombo.SelectedItem as string;

        _viewConfig.Sort.Clear();
        if (!string.IsNullOrEmpty(field) && field != "(不排序)")
        {
            _viewConfig.Sort.Add(new SortConfig
            {
                Field = field,
                Order = orderText == "降序" ? "desc" : "asc"
            });
        }

        RefreshViewData();
        SaveConfig();
    }

    /// <summary>按当前配置重新计算视图数据并刷新看板</summary>
    private void RefreshViewData()
    {
        if (_dataTable == null || _viewConfig == null) return;
        _viewData = new ViewEngine().Apply(_dataTable, _viewConfig);
        BuildColumns();
    }

    /// <summary>保存视图配置</summary>
    private void SaveConfig()
    {
        if (_configFile == null || _excelAdapter == null || _viewConfig == null) return;
        try
        {
            var path = _excelAdapter.GetActiveWorkbookPath();
            new ViewConfigManager().Save(path, _configFile, _configFile.TableName);
        }
        catch (Exception ex)
        {
            AddInLog.Write("KanbanView.SaveConfig.Error", ex.ToString());
        }
    }
}
