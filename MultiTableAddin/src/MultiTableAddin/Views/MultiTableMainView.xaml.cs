using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using UserControl = System.Windows.Controls.UserControl;
using MessageBox = System.Windows.MessageBox;
using MessageBoxResult = System.Windows.MessageBoxResult;
using MenuItem = System.Windows.Controls.MenuItem;
using MultiTableAddin.Core;

namespace MultiTableAddin.Views;

public partial class MultiTableMainView : UserControl, IRequestClose
{
    private readonly ExcelAdapter _excelAdapter = new();
    private readonly ViewEngine _viewEngine = new();
    private readonly ViewConfigManager _configManager = new();

    private DataTableModel? _dataTable;
    private ViewConfigFile? _configFile;
    private ViewConfig? _currentView;
    private readonly Dictionary<string, UserControl> _viewCache = new();

    private static readonly Dictionary<ViewType, string> ViewIcons = new()
    {
        { ViewType.Table, "\ue8a9" },
        { ViewType.Form, "\ue8b7" },
        { ViewType.Kanban, "\ue8a3" },
        { ViewType.Gallery, "\ue8b5" },
        { ViewType.Calendar, "\ue8b8" },
        { ViewType.Gantt, "\ue8ba" },
        { ViewType.Dashboard, "\ue8c0" },
        { ViewType.Chart, "\ue8c1" }
    };

    private static readonly Dictionary<ViewType, string> ViewLabels = new()
    {
        { ViewType.Table, "表格" },
        { ViewType.Form, "表单" },
        { ViewType.Kanban, "看板" },
        { ViewType.Gallery, "画册" },
        { ViewType.Calendar, "日历" },
        { ViewType.Gantt, "甘特" },
        { ViewType.Dashboard, "仪表盘" },
        { ViewType.Chart, "图表" }
    };

    public event EventHandler? RequestClose;

    public MultiTableMainView()
    {
        InitializeComponent();
        VersionText.Text = AppVersion.DisplayText;
    }

    private void MultiTableMainView_Loaded(object sender, RoutedEventArgs e)
    {
        TryAutoLoadData();
    }

    private void MultiTableMainView_Unloaded(object sender, RoutedEventArgs e)
    {
        _excelAdapter.Dispose();
    }

    /// <summary>尝试自动加载活动工作簿中的 ListObject</summary>
    private void TryAutoLoadData()
    {
        try
        {
            var names = _excelAdapter.GetListObjectNames();
            if (names.Count == 0)
            {
                HandyControl.Controls.Growl.InfoGlobal("未找到超级表(ListObject)，请先将数据区域转换为超级表(Ctrl+T)后再加载。");
                return;
            }

            // 默认加载第一个 ListObject
            string sheetName = _excelAdapter.GetSheetNames().FirstOrDefault() ?? string.Empty;
            LoadData(sheetName, names[0]);
        }
        catch (Exception ex)
        {
            AddInLog.Write("MultiTableMainView.TryAutoLoad.Error", ex.ToString());
            HandyControl.Controls.Growl.WarningGlobal("自动加载数据失败: " + ex.Message);
        }
    }

    /// <summary>加载指定工作表和表名的数据</summary>
    public void LoadData(string sheetName, string tableName, bool switchToDefault = true)
    {
        try
        {
            _dataTable = _excelAdapter.ReadListObject(sheetName, tableName);

            if (_dataTable.Rows.Count == 0)
            {
                HandyControl.Controls.Growl.WarningGlobal("表 '" + tableName + "' 没有数据行。");
                return;
            }

            // 加载或创建视图配置（按超级表名称隔离，避免不同表共用同一套视图）
            string wbPath = _excelAdapter.GetActiveWorkbookPath();
            _configFile = _configManager.Load(wbPath, tableName);

            // 如果没有视图配置，创建默认配置；否则同步字段变化
            if (_configFile.Views.Count == 0)
            {
                _configFile = _configManager.CreateDefaultConfig(_dataTable);
            }
            else
            {
                ViewConfigManager.SyncFields(_configFile, _dataTable);
            }

            _configManager.Save(wbPath, _configFile, tableName);

            // 同步字段信息
            _configFile.Fields = _dataTable.Fields;
            _configFile.SourceSheet = sheetName;
            _configFile.TableName = tableName;

            // 更新界面
            SourceInfoText.Text = $"{sheetName} / {tableName} ({_dataTable.Rows.Count} 行)";
            BuildViewList();

            // 默认显示第一个视图（刷新时不自动切换）
            if (switchToDefault && _configFile.Views.Count > 0)
            {
                SwitchView(_configFile.Views[0]);
            }

            // 显示操作按钮
            BtnRefresh.Visibility = Visibility.Visible;
            BtnSave.Visibility = Visibility.Visible;
            BtnAddView.Visibility = Visibility.Visible;
            BtnFieldConfig.Visibility = Visibility.Visible;

            AddInLog.Write("MultiTableMainView.LoadData", $"Sheet={sheetName}, Table={tableName}, Rows={_dataTable.Rows.Count}");
        }
        catch (Exception ex)
        {
            AddInLog.Write("MultiTableMainView.LoadData.Error", ex.ToString());
            HandyControl.Controls.Growl.ErrorGlobal("加载数据失败: " + ex.Message);
        }
    }

    /// <summary>构建左侧视图列表</summary>
    private void BuildViewList()
    {
        ViewListPanel.Children.Clear();
        if (_configFile == null) return;

        foreach (var view in _configFile.Views)
        {
            var textBlock = new TextBlock
            {
                Text = view.ViewName,
                FontSize = 12,
                TextTrimming = TextTrimming.CharacterEllipsis,
                TextWrapping = TextWrapping.NoWrap,
                VerticalAlignment = VerticalAlignment.Center
            };

            var btn = new Button
            {
                Content = textBlock,
                Tag = view,
                Style = (Style)FindResource("SidebarButton")
            };

            if (view == _currentView)
            {
                btn.Style = (Style)FindResource("SidebarButtonActive");
            }

            btn.Click += ViewButton_Click;
            btn.ToolTip = "左键切换视图 · 右键重命名/删除";

            var menu = new ContextMenu();
            var renameItem = new MenuItem { Header = "重命名" };
            renameItem.Click += (_, _) => RenameView(view);
            var deleteItem = new MenuItem { Header = "删除" };
            deleteItem.Click += (_, _) => DeleteView(view);
            menu.Items.Add(renameItem);
            menu.Items.Add(deleteItem);
            btn.ContextMenu = menu;

            ViewListPanel.Children.Add(btn);
        }
    }

    /// <summary>重命名指定视图</summary>
    private void RenameView(ViewConfig view)
    {
        if (_configFile == null) return;
        var dlg = new InputPromptWindow("重命名视图", "请输入视图名称：", view.ViewName);
        dlg.Owner = Window.GetWindow(this);
        if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.ResultText))
        {
            view.ViewName = dlg.ResultText.Trim();
            try
            {
                _configManager.Save(_excelAdapter.GetActiveWorkbookPath(), _configFile, _configFile.TableName);
            }
            catch (Exception ex)
            {
                HandyControl.Controls.Growl.ErrorGlobal("保存失败: " + ex.Message);
            }
            BuildViewList();
            if (_currentView == view)
            {
                CurrentViewName.Text = view.ViewName;
            }
        }
    }

    /// <summary>删除指定视图</summary>
    private void DeleteView(ViewConfig view)
    {
        if (_configFile == null) return;
        var result = MessageBox.Show($"确定删除视图「{view.ViewName}」吗？此操作不可撤销。",
            "删除视图", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        _configFile.Views.Remove(view);

        if (_currentView == view)
        {
            _currentView = null;
            ViewContentArea.Children.Clear();
            _viewCache.Clear();
            CurrentViewName.Text = "请加载数据";
            FilterInfo.Text = "";
        }

        try
        {
            _configManager.Save(_excelAdapter.GetActiveWorkbookPath(), _configFile, _configFile.TableName);
        }
        catch (Exception ex)
        {
            HandyControl.Controls.Growl.ErrorGlobal("保存失败: " + ex.Message);
        }

        BuildViewList();
    }

    private void ViewButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is ViewConfig view)
        {
            SwitchView(view);
        }
    }

    /// <summary>切换到指定视图</summary>
    private void SwitchView(ViewConfig view)
    {
        if (_dataTable == null) return;

        _currentView = view;
        CurrentViewName.Text = view.ViewName;
        FilterInfo.Text = string.IsNullOrWhiteSpace(view.Filter) ? "" : $"筛选: {view.Filter}";

        // 刷新侧边栏高亮
        BuildViewList();

        // 清空内容区域
        ViewContentArea.Children.Clear();
        _viewCache.Clear();

        // 创建并显示对应视图
        var viewControl = CreateViewControl(view);
        if (viewControl != null)
        {
            ViewContentArea.Children.Add(viewControl);
        }
    }

    /// <summary>根据视图类型创建视图控件</summary>
    private UserControl? CreateViewControl(ViewConfig view)
    {
        if (_dataTable == null) return null;

        var viewData = _viewEngine.Apply(_dataTable, view);

        UserControl? control = view.ViewType switch
        {
            ViewType.Table => new TableView(),
            ViewType.Form => new FormView(),
            ViewType.Kanban => new KanbanView(),
            ViewType.Gallery => new GalleryView(),
            ViewType.Calendar => new CalendarView(),
            ViewType.Gantt => new GanttView(),
            ViewType.Dashboard => new DashboardView(),
            ViewType.Chart => new ChartView(),
            _ => null
        };

        if (control == null) return null;

        // 传递数据给视图
        if (control is ITableView tv)
        {
            tv.Initialize(_dataTable, view, viewData, _excelAdapter);
        }

        // 传递配置文件给需要字段配置的视图
        if (control is IConfigAware ca && _configFile != null)
        {
            ca.SetConfigFile(_configFile);
        }

        return control;
    }

    private void OnLoadDataClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var sources = _excelAdapter.GetTableSources();
            if (sources.Count == 0)
            {
                HandyControl.Controls.Growl.WarningGlobal("未找到超级表(ListObject)。请先将数据区域转换为超级表(Ctrl+T)。");
                return;
            }

            var dlg = new DataSourcePickerDialog(sources);
            dlg.Owner = Window.GetWindow(this);
            if (dlg.ShowDialog() == true && dlg.Selected != null)
            {
                LoadData(dlg.Selected.SheetName, dlg.Selected.TableName);
            }
        }
        catch (Exception ex)
        {
            HandyControl.Controls.Growl.ErrorGlobal("加载数据失败: " + ex.Message);
        }
    }

    private void OnRefreshClick(object sender, RoutedEventArgs e)
    {
        if (_dataTable == null || _currentView == null) return;
        if (_configFile != null)
        {
            var viewId = _currentView.ViewId;
            LoadData(_configFile.SourceSheet, _configFile.TableName, switchToDefault: false);

            // 刷新后尽量回到刷新前的视图
            var restored = _configFile.Views.FirstOrDefault(v => v.ViewId == viewId);
            if (restored != null)
            {
                SwitchView(restored);
            }
            else if (_configFile.Views.Count > 0)
            {
                SwitchView(_configFile.Views[0]);
            }

            HandyControl.Controls.Growl.SuccessGlobal("数据已刷新。");
        }
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        if (_configFile == null) return;

        try
        {
            string wbPath = _excelAdapter.GetActiveWorkbookPath();
            _configManager.Save(wbPath, _configFile, _configFile.TableName);
            HandyControl.Controls.Growl.SuccessGlobal("视图配置已保存。");
        }
        catch (Exception ex)
        {
            HandyControl.Controls.Growl.ErrorGlobal("保存失败: " + ex.Message);
        }
    }

    private void OnAddViewClick(object sender, RoutedEventArgs e)
    {
        if (_configFile == null || _dataTable == null)
        {
            HandyControl.Controls.Growl.WarningGlobal("请先加载数据。");
            return;
        }

        // 选择要创建的视图类型（含甘特等全部类型）
        var dlg = new AddViewDialog();
        dlg.Owner = Window.GetWindow(this);
        if (dlg.ShowDialog() != true) return;

        var newView = CreateViewConfig(dlg.SelectedType);
        _configFile.Views.Add(newView);
        BuildViewList();
        SwitchView(newView);
    }

    /// <summary>为指定视图类型创建一个带合理默认值的视图配置</summary>
    private ViewConfig CreateViewConfig(ViewType type)
    {
        var allNames = _dataTable!.FieldNames;
        var firstName = allNames.Count > 0 ? allNames[0] : string.Empty;

        var cfg = new ViewConfig
        {
            ViewId = "view_" + Guid.NewGuid().ToString("N").Substring(0, 8),
            ViewType = type,
            ViewName = ViewConfigManager.DefaultViewName(type),
            VisibleFields = new List<string>(allNames),
            Sort = new List<SortConfig>()
        };

        if (type == ViewType.Gantt)
        {
            var dateFields = _dataTable.Fields.FindAll(f => FieldTypeHelper.IsTemporal(f.Type));
            cfg.GanttConfig = new GanttConfig
            {
                StartField = dateFields.Count > 0 ? dateFields[0].Name : string.Empty,
                EndField = dateFields.Count > 1 ? dateFields[1].Name
                    : (dateFields.Count > 0 ? dateFields[0].Name : string.Empty),
                LabelField = firstName
            };
        }

        return cfg;
    }

    private void OnFieldConfigClick(object sender, RoutedEventArgs e)
    {
        if (_configFile == null || _dataTable == null)
        {
            HandyControl.Controls.Growl.WarningGlobal("请先加载数据。");
            return;
        }

        var dialog = new FieldConfigDialog(_dataTable, _configFile);
        dialog.Owner = Window.GetWindow(this);
        dialog.ShowDialog();

        // 如果用户保存了配置，刷新当前视图
        if (dialog.ConfigChanged)
        {
            if (_currentView != null)
            {
                SwitchView(_currentView);
            }
        }
    }
}

/// <summary>视图控件接口，用于传递数据</summary>
public interface ITableView
{
    void Initialize(DataTableModel dataTable, ViewConfig viewConfig, ViewDataSet viewData, ExcelAdapter excelAdapter);
}

/// <summary>需要访问配置文件的视图接口（如表单视图需要字段覆盖配置）</summary>
public interface IConfigAware
{
    void SetConfigFile(ViewConfigFile configFile);
}
