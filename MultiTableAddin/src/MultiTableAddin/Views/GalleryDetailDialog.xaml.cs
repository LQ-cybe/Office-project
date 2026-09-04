using System.Linq;
using System.Windows;
using System.Windows.Input;
using MultiTableAddin.Core;
using MultiTableAddin.Views.Controls;

namespace MultiTableAddin.Views;

/// <summary>
/// 画册记录详情对话框：复用 RecordEditor 提供人性化的查看 / 编辑界面，
/// 可在表单中直接修改并（由调用方）写回 Excel，也支持定位到表格对应行。
/// </summary>
public partial class GalleryDetailDialog : Window
{
    private readonly DataTableModel _table;
    private readonly DataRowModel _row;
    private readonly ViewConfigFile? _config;
    private readonly ExcelAdapter? _excel;

    public GalleryDetailDialog(DataTableModel dataTable, DataRowModel row,
        ViewConfigFile? configFile, ExcelAdapter? excelAdapter = null)
    {
        InitializeComponent();

        _table = dataTable;
        _row = row;
        _config = configFile;
        _excel = excelAdapter;

        Title = "记录详情 · 第 " + row.RowIndex + " 行";
        Editor.ReadOnlyMode = false;
        Editor.HideAffixes = true;          // 不显示金额/百分比的 ¥、% 符号
        Editor.UniformControlWidth = 250;   // 所有录入控件宽度保持一致
        Editor.Load(_table, _config, _row, null);

        bool canLocate = _excel != null && !string.IsNullOrEmpty(_table.SheetName);
        BtnLocate.Visibility = canLocate ? Visibility.Visible : Visibility.Collapsed;

        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        };
    }

    /// <summary>返回相对原值发生变化的字段（供调用方写回 Excel）</summary>
    public Dictionary<string, object?> GetChanges() => Editor.GetChanges();

    private void OnLocateClick(object sender, RoutedEventArgs e)
    {
        try
        {
            _excel?.SelectRow(_table.SheetName, _table.TableName, _row.RowIndex);
        }
        catch (Exception ex)
        {
            HandyControl.Controls.Growl.WarningGlobal("定位失败: " + ex.Message);
        }
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        var errors = Editor.Validate();
        if (errors.Count > 0)
        {
            HandyControl.Controls.Growl.WarningGlobal("校验未通过：\n" + string.Join("\n", errors.Select(kv => $"• {kv.Key}: {kv.Value}")));
            return;
        }
        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
