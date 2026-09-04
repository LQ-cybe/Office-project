using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using UserControl = System.Windows.Controls.UserControl;
using System.Windows.Media.Imaging;
using MultiTableAddin.Core;

namespace MultiTableAddin.Views;

public partial class GalleryView : UserControl, ITableView, IConfigAware
{
    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;
    private ViewConfigFile? _configFile;
    private List<DataRowModel> _allRows = new();
    private Border? _selectedCard;

    public GalleryView()
    {
        InitializeComponent();
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;
        BuildToolbar();
        BuildGallery();
    }

    /// <summary>构建顶部工具栏选项</summary>
    private void BuildToolbar()
    {
        if (_dataTable == null || _viewConfig == null) return;

        var fieldNames = _dataTable.Fields.Select(f => f.Name).ToList();
        var optionalFields = new List<string> { "(无)" }.Concat(fieldNames).ToList();

        TitleFieldCombo.ItemsSource = fieldNames;
        ImageFieldCombo.ItemsSource = optionalFields;

        TitleFieldCombo.SelectionChanged -= OnTitleFieldChanged;
        ImageFieldCombo.SelectionChanged -= OnImageFieldChanged;

        var meta = _viewConfig.CardMeta ??= new CardMeta();
        TitleFieldCombo.SelectedItem = fieldNames.Contains(meta.Title) ? meta.Title : fieldNames.FirstOrDefault();
        ImageFieldCombo.SelectedItem = string.IsNullOrEmpty(meta.Image) ? "(无)" : meta.Image;

        TitleFieldCombo.SelectionChanged += OnTitleFieldChanged;
        ImageFieldCombo.SelectionChanged += OnImageFieldChanged;

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

    public void SetConfigFile(ViewConfigFile configFile)
    {
        _configFile = configFile;
    }

    private void BuildGallery()
    {
        if (_viewData == null || _viewConfig?.CardMeta == null || _dataTable == null) return;

        _allRows = _viewData.Groups.SelectMany(g => g.Rows).ToList();
        GalleryItems.Items.Clear();

        string titleField = _viewConfig.CardMeta.Title;
        string imageField = _viewConfig.CardMeta.Image;
        var descFields = _viewConfig.CardMeta.Description ?? new List<string>();

        // 动态计算卡片宽度：基于最长字段名
        double maxLabelWidth = 0;
        var allFields = new List<string> { titleField };
        allFields.AddRange(descFields);
        foreach (var f in allFields.Where(f => !string.IsNullOrEmpty(f)))
        {
            var formatted = new FormattedText(
                $"{f}: 测试值",
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Microsoft YaHei"),
                12, Brushes.Black, 1.0);
            if (formatted.Width > maxLabelWidth) maxLabelWidth = formatted.Width;
        }
        double cardWidth = Math.Max(220, maxLabelWidth + 40);
        cardWidth = Math.Min(cardWidth, 400); // 上限 400

        foreach (var row in _allRows)
        {
            var card = CreateCard(row, titleField, imageField, descFields, cardWidth);
            GalleryItems.Items.Add(card);
        }
    }

    private Border CreateCard(DataRowModel row, string titleField, string imageField,
        List<string> descFields, double cardWidth)
    {
        var card = new Border
        {
            Style = (Style)FindResource("GalleryCard"),
            Width = cardWidth,
            Cursor = System.Windows.Input.Cursors.Hand,
            Tag = row
        };

        var cardPanel = new StackPanel();

        // 图片区域
        if (!string.IsNullOrEmpty(imageField))
        {
            var imagePath = row.GetValue(imageField)?.ToString();
            var image = CreateImage(imagePath, cardWidth - 2);
            cardPanel.Children.Add(image);
        }

        var contentPanel = new StackPanel { Margin = new Thickness(12, 8, 12, 12) };

        // 标题
        if (!string.IsNullOrEmpty(titleField))
        {
            contentPanel.Children.Add(new TextBlock
            {
                Text = row.GetValue(titleField)?.ToString() ?? "(空)",
                FontSize = 14,
                FontWeight = FontWeights.Medium,
                Foreground = TryFindResource("PrimaryTextBrush") as Brush,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(0, 0, 0, 6)
            });
        }

        // 描述字段
        foreach (var descField in descFields)
        {
            var val = row.GetValue(descField);
            if (val == null) continue;
            contentPanel.Children.Add(new TextBlock
            {
                Text = $"{descField}: {val}",
                FontSize = 12,
                Foreground = TryFindResource("SecondaryTextBrush") as Brush,
                Margin = new Thickness(0, 2, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            });
        }

        cardPanel.Children.Add(contentPanel);
        card.Child = cardPanel;

        // 点击事件 — 打开详情对话框
        card.MouseLeftButtonUp += (s, e) => OpenDetailDialog(row);

        return card;
    }

    /// <summary>打开记录详情对话框</summary>
    private void OpenDetailDialog(DataRowModel row)
    {
        if (_dataTable == null) return;

        // 高亮当前选中的画册卡片（与看板视图一致）
        HighlightCard(row);

        var dialog = new GalleryDetailDialog(_dataTable, row, _configFile, _excelAdapter);
        dialog.Owner = Window.GetWindow(this);

        if (dialog.ShowDialog() == true && _excelAdapter != null)
        {
            // 保存修改到 Excel
            var changes = dialog.GetChanges();
            foreach (var kv in changes)
            {
                _excelAdapter.UpdateCell(_dataTable.SheetName, _dataTable.TableName, row.RowIndex, kv.Key, kv.Value);
                row.SetValue(kv.Key, kv.Value);
            }
            _dataTable.IsDirty = true;
            HandyControl.Controls.Growl.SuccessGlobal("修改已保存到 Excel。");

            // 刷新画册
            BuildGallery();
        }

        // 关闭详情后恢复卡片样式
        ClearCardHighlight();
    }

    /// <summary>高亮与指定行对应的画册卡片</summary>
    private void HighlightCard(DataRowModel row)
    {
        var card = GalleryItems.Items.OfType<Border>().FirstOrDefault(b => b.Tag == row);
        if (card == null) return;

        if (_selectedCard != null && _selectedCard != card)
            ResetCardStyle(_selectedCard);

        _selectedCard = card;
        card.BorderBrush = TryFindResource("PrimaryBrush") as Brush;
        card.BorderThickness = new Thickness(2);
        card.Background = TryFindResource("LightPrimaryBrush") as Brush ?? TryFindResource("RegionBrush") as Brush;
    }

    private static void ResetCardStyle(Border card)
    {
        card.BorderBrush = card.TryFindResource("BorderBrush") as Brush;
        card.BorderThickness = new Thickness(1);
        card.Background = card.TryFindResource("RegionBrush") as Brush;
    }

    private void ClearCardHighlight()
    {
        if (_selectedCard != null) ResetCardStyle(_selectedCard);
        _selectedCard = null;
    }

    /// <summary>标题字段切换</summary>
    private void OnTitleFieldChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_viewConfig?.CardMeta == null || _dataTable == null) return;
        var field = TitleFieldCombo.SelectedItem as string;
        if (!string.IsNullOrEmpty(field))
        {
            _viewConfig.CardMeta.Title = field;
            SaveConfig();
            BuildGallery();
        }
    }

    /// <summary>图片字段切换</summary>
    private void OnImageFieldChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_viewConfig?.CardMeta == null || _dataTable == null) return;
        var field = ImageFieldCombo.SelectedItem as string;
        _viewConfig.CardMeta.Image = field == "(无)" ? string.Empty : field ?? string.Empty;
        SaveConfig();
        BuildGallery();
    }

    /// <summary>字段显隐：选择描述字段</summary>
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
            SaveConfig();
            BuildGallery();
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

    /// <summary>按当前配置重新计算视图数据并刷新画册</summary>
    private void RefreshViewData()
    {
        if (_dataTable == null || _viewConfig == null) return;
        _viewData = new ViewEngine().Apply(_dataTable, _viewConfig);
        BuildGallery();
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
            AddInLog.Write("GalleryView.SaveConfig.Error", ex.ToString());
        }
    }

    private System.Windows.Controls.Image CreateImage(string? imagePath, double targetWidth)
    {
        double imgHeight = Math.Min(130, targetWidth * 0.6);
        var image = new System.Windows.Controls.Image
        {
            Width = targetWidth,
            Height = imgHeight,
            Stretch = Stretch.UniformToFill,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0),
            Clip = new RectangleGeometry(new Rect(0, 0, targetWidth, imgHeight), 8, 8)
        };

        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
        {
            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                bitmap.DecodePixelWidth = (int)targetWidth;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                image.Source = bitmap;
            }
            catch
            {
                image.Source = CreatePlaceholderImage(targetWidth, imgHeight);
            }
        }
        else
        {
            image.Source = CreatePlaceholderImage(targetWidth, imgHeight);
        }

        return image;
    }

    private static DrawingImage CreatePlaceholderImage(double w, double h)
    {
        var dv = new DrawingVisual();
        using (var dc = dv.RenderOpen())
        {
            dc.DrawRectangle(new SolidColorBrush(Color.FromRgb(240, 240, 240)), null, new Rect(0, 0, w, h));
            var text = new FormattedText("无图片",
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Microsoft YaHei"),
                14, new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                1.0);
            dc.DrawText(text, new Point(w / 2 - 20, h / 2 - 10));
        }
        return new DrawingImage(dv.Drawing);
    }
}
