using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using UserControl = System.Windows.Controls.UserControl;
using FontFamily = System.Windows.Media.FontFamily;
using TextBox = System.Windows.Controls.TextBox;
using MultiTableAddin.Core;

namespace MultiTableAddin.Views;

public partial class TableView : UserControl, ITableView
{
    private DataTableModel? _dataTable;
    private ViewConfig? _viewConfig;
    private ViewDataSet? _viewData;
    private ExcelAdapter? _excelAdapter;

    public TableView()
    {
        InitializeComponent();
        MainDataGrid.LoadingRow += (_, e) => e.Row.Header = (e.Row.GetIndex() + 1).ToString();
    }

    public void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter)
    {
        _dataTable = dataTable;
        _viewConfig = viewConfig;
        _viewData = viewData;
        _excelAdapter = excelAdapter;

        BuildColumns();
        LoadData();
    }

    private void BuildColumns()
    {
        MainDataGrid.Columns.Clear();
        if (_viewConfig == null || _dataTable == null) return;

        var tableCfg = _viewConfig.TableConfig ?? new TableViewConfig();
        bool showRowNumber = tableCfg.ShowRowNumber;

        MainDataGrid.HeadersVisibility = showRowNumber
            ? DataGridHeadersVisibility.All
            : DataGridHeadersVisibility.Column;
        MainDataGrid.RowHeaderWidth = showRowNumber ? 44 : 0;

        var visibleFields = _viewConfig.VisibleFields.Count > 0
            ? _viewConfig.VisibleFields
            : _dataTable.Fields.ConvertAll(f => f.Name);

        int actualColumns = 0;
        foreach (var fieldName in visibleFields)
        {
            var field = _dataTable.Fields.Find(f => f.Name == fieldName);
            if (field == null) continue;

            actualColumns++;
            var column = CreateTemplateColumn(field, tableCfg);
            MainDataGrid.Columns.Add(column);
        }

        ColCountText.Text = $"{actualColumns} 列";
    }

    /// <summary>创建模板列：显示用 Excel 实际文本，编辑用原始值；数字右对齐、文本左对齐；列宽按内容自适应</summary>
    private DataGridTemplateColumn CreateTemplateColumn(FieldSchema field, TableViewConfig cfg)
    {
        bool rightAlign = FieldTypeHelper.IsNumeric(field.Type);

        // 显示模板：TextBlock 绑定到 Row.DisplayTexts[fieldName]
        var displayFactory = new FrameworkElementFactory(typeof(TextBlock));
        displayFactory.SetBinding(TextBlock.TextProperty, new Binding($"Row.DisplayTexts[{field.Name}]"));
        displayFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
        displayFactory.SetValue(TextBlock.HorizontalAlignmentProperty,
            rightAlign ? HorizontalAlignment.Right : HorizontalAlignment.Left);
        displayFactory.SetValue(TextBlock.TextAlignmentProperty,
            rightAlign ? TextAlignment.Right : TextAlignment.Left);
        displayFactory.SetValue(TextBlock.PaddingProperty, new Thickness(8, 0, 8, 0));
        displayFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);

        var cellTemplate = new DataTemplate { VisualTree = displayFactory };

        // 编辑模板：TextBox 绑定到 RawValues[field.Name]，按字段类型解析
        var editFactory = new FrameworkElementFactory(typeof(TextBox));
        editFactory.SetBinding(TextBox.TextProperty,
            new Binding($"RawValues[{field.Name}]") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
        editFactory.SetValue(TextBox.VerticalAlignmentProperty, VerticalAlignment.Center);
        editFactory.SetValue(TextBox.HorizontalAlignmentProperty,
            rightAlign ? HorizontalAlignment.Right : HorizontalAlignment.Left);
        editFactory.SetValue(TextBox.PaddingProperty, new Thickness(6, 0, 6, 0));
        editFactory.SetValue(TextBox.BorderThicknessProperty, new Thickness(0));

        var editingTemplate = new DataTemplate { VisualTree = editFactory };

        return new DataGridTemplateColumn
        {
            Header = field.Name,
            CellTemplate = cellTemplate,
            CellEditingTemplate = editingTemplate,
            SortMemberPath = $"Row.Values[{field.Name}]",
            Width = DataGridLength.Auto,
            MinWidth = cfg.MinColumnWidth,
            MaxWidth = cfg.MaxColumnWidth
        };
    }

    private void LoadData()
    {
        if (_viewData == null) return;

        // 将 DataRowModel 转为可编辑的动态字典列表
        var items = new ObservableCollection<EditableDataRow>();
        foreach (var group in _viewData.Groups)
        {
            foreach (var row in group.Rows)
            {
                items.Add(new EditableDataRow(row, _dataTable!, _excelAdapter!));
            }
        }

        MainDataGrid.ItemsSource = items;
        RowCountText.Text = $"{items.Count} 行";

        // 列宽设为 Auto，由 WPF 按表头与单元格内容自适应；MaxWidth 限制超长文本
        MainDataGrid.UpdateLayout();
    }
}

/// <summary>可编辑的数据行，支持双向绑定</summary>
public class EditableDataRow : INotifyPropertyChanged
{
    private readonly DataRowModel _row;
    private readonly DataTableModel _table;
    private readonly ExcelAdapter _adapter;

    public EditableDataRow(DataRowModel row, DataTableModel table, ExcelAdapter adapter)
    {
        _row = row;
        _table = table;
        _adapter = adapter;
        RawValues = new RawValueProxy(this);
    }

    /// <summary>原始数据行，用于表格视图显示 Excel 实际文本</summary>
    public DataRowModel Row => _row;

    /// <summary>编辑时访问的原始值包装器，写入时按字段类型解析</summary>
    public RawValueProxy RawValues { get; }

    /// <summary>索引器绑定，支持 DataGrid Binding [字段名]</summary>
    public object? this[string fieldName]
    {
        get => _row.GetValue(fieldName);
        set => SetFieldValue(fieldName, value);
    }

    /// <summary>按字段类型解析并写入，同时同步 Excel 与显示文本</summary>
    private void SetFieldValue(string fieldName, object? value)
    {
        var field = _table.FindField(fieldName);
        object? parsed = value;
        if (value is string s && field != null)
        {
            parsed = ValueFormatter.ParseInput(s, field.Type);
        }

        if (!Equals(_row.GetValue(fieldName), parsed))
        {
            _row.SetValue(fieldName, parsed);
            // 显示文本按字段类型重新格式化（刷新后会从 Excel Text 重新读取精确格式）
            _row.DisplayTexts[fieldName] = field != null
                ? ValueFormatter.ToDisplayText(parsed, field.Type)
                : ValueFormatter.ToDisplayText(parsed);
            // 回写到 Excel
            _adapter.UpdateCell(_table.SheetName, _table.TableName, _row.RowIndex, fieldName, parsed);
            _table.IsDirty = true;
            OnPropertyChanged($"Item[{fieldName}]");
            OnPropertyChanged($"RawValues[{fieldName}]");
            OnPropertyChanged($"Row.DisplayTexts[{fieldName}]");
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged(string propName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));

    /// <summary>编辑模板使用的原始值代理，set 时按字段类型解析</summary>
    public class RawValueProxy : INotifyPropertyChanged
    {
        private readonly EditableDataRow _parent;
        public RawValueProxy(EditableDataRow parent) => _parent = parent;

        public object? this[string fieldName]
        {
            get => _parent._row.GetValue(fieldName);
            set => _parent.SetFieldValue(fieldName, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        internal void Notify(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
