using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using MultiTableAddin.Core;
using MultiTableAddin.Views.Controls;

namespace MultiTableAddin.Views;

/// <summary>图表配置弹窗：左侧实时预览，右侧参数面板</summary>
public partial class ChartConfigDialog : Window
{
    private readonly DataTableModel _table;
    private readonly ViewConfigFile? _config;
    private readonly IReadOnlyList<DataRowModel> _rows;
    private readonly ChartConfigEditor _editor = new();

    /// <summary>编辑副本，点击确定后由调用方取回</summary>
    public ChartConfig Result { get; }

    public ChartConfigDialog(DataTableModel table, ViewConfigFile? config,
        ChartConfig source, IReadOnlyList<DataRowModel> rows)
    {
        InitializeComponent();

        _table = table;
        _config = config;
        _rows = rows ?? Array.Empty<DataRowModel>();
        Result = source.Clone();

        EditorHost.Content = _editor;
        _editor.Changed += (_, _) => RefreshPreview();
        _editor.Load(_table, _config, Result);

        RefreshPreview();

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        };
    }

    private void RefreshPreview()
    {
        try
        {
            var ds = ChartDataBuilder.Build(_rows, Result);
            Preview.ChartType = Result.Type;
            Preview.GaugeTarget = Result.GaugeTarget;
            Preview.Data = ds;

            PreviewTitle.Text = string.IsNullOrWhiteSpace(Result.Title) ? "预览" : Result.Title;
            PreviewMessage.Text = ds.Message;
        }
        catch (Exception ex)
        {
            PreviewMessage.Text = "预览失败: " + ex.Message;
        }
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(Result.Id))
            Result.Id = ViewConfig.NewId("chart");
        if (string.IsNullOrWhiteSpace(Result.Title))
            Result.Title = "未命名图表";

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
