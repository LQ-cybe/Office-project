# Excel 插件方案可行性分析 — 多维表功能以插件形式实现

> 文档版本：v1.0  
> 日期：2026-08-07  
> 定位：分析以 Excel 插件形式（VSTO / COM Add-in / Web Add-in）实现多维表自定义窗口界面的可行性

---

## 一、结论先行

**可行，但有明确的天花板和权衡。**

以 VSTO/COM Add-in 形式在 Excel 中实现多维表功能，技术上完全可行，且你有 VBE2021 项目的 VSTO 开发经验，技术栈高度匹配。但存在以下核心权衡：

| 维度 | 独立 EXE 方案 | Excel 插件方案 |
|------|-------------|---------------|
| 开发效率 | 中（需自建文件解析） | **中高**（复用 Excel 数据引擎） |
| 数据存储 | 需自建文件 IO | **直接用 Excel 单元格** |
| 视图配置存储 | 外置 JSON | 可嵌入 xlsx 自定义 XML 部件 |
| 分发部署 | 单 exe，无依赖 | ⚠️ 需安装 Office 运行时 + VSTO 部署 |
| 宿主依赖 | 无 | **绑定 MS Excel** |
| WPS 兼容 | N/A | ❌ 不兼容 |
| UI 自由度 | 完全自定义 | 受 TaskPane 尺寸限制 |
| 用户接受度 | 需学习新软件 | **用户已在 Excel 中，零迁移成本** |

**适用场景**：
- 用户日常工作深度依赖 Excel，不愿意切换到独立软件
- 内部团队使用，统一安装 MS Office 32/64 位
- 需要 Excel 原生公式、条件格式与多维表视图共存

**不适用场景**：
- 需要分发给外部客户，Office 版本不可控
- 需要 WPS 兼容
- 需要完全脱离 Office 独立运行

---

## 二、技术方案详解

### 2.1 方案总览：三种 Excel 插件技术

| 技术方案 | 技术栈 | UI 能力 | 分发难度 | WPS 兼容 |
|---------|--------|---------|---------|---------|
| **VSTO Add-in** | C# + .NET Framework / .NET 6-8 | Custom TaskPane + WPF/WinForms | ClickOnce 部署 | ❌ |
| **COM Add-in (Shared Add-in)** | C# / VB.NET + Office Interop | Custom TaskPane | 注册表 + 安装包 | ❌ |
| **Office Web Add-in** | HTML/JS + Office.js | Web TaskPane | Office Store / 旁加载 | ⚠️ 部分 |

### 2.2 推荐方案：VSTO + WebView2 + Vue3 前端

结合你的技术背景（VB.NET/COM、VSTO 开发经验）和多维表视图需求，推荐以下技术组合：

```
VSTO Excel Add-in (.NET 8, C#)
├── Custom Task Pane（自定义任务窗格）
│   └── WebView2 控件
│       └── Vue3 前端应用（6大视图 + 交互组件）
│
├── C# 后端服务层
│   ├── ExcelAdapter：读写 ListObject（超级表）/ Range
│   ├── ViewConfigManager：视图配置管理（自定义 XML 部件 / 外置 JSON）
│   ├── ViewEngine：过滤、排序、分组引擎
│   ├── DataSyncManager：双向数据同步
│   └── EventBridge：监听 Excel 工作表变更事件
│
├── 通信桥接层
│   └── WebView2 PostMessage / WebMessageReceived
│
└── 视图配置存储
    ├── 首选：工作簿自定义 XML 部件（嵌入 xlsx 内部）
    └── 备选：外置 {文件名}.multiview.config.json
```

### 2.3 备选方案：VSTO + 纯 WPF（无 WebView 依赖）

如果不想引入 WebView2 运行时依赖，可以使用纯 WPF 实现全部视图：

```
VSTO Excel Add-in (.NET 8, C#)
├── Custom Task Pane
│   └── WPF UserControl（ElementHost 承载）
│       ├── 表格视图：DataGrid（虚拟化）
│       ├── 看板视图：ItemsControl + 拖拽（自研）
│       ├── 画册视图：WrapPanel（需自研虚拟化）
│       ├── 日历视图：自研日历面板
│       ├── 甘特视图：自研时间轴 + 进度条
│       └── 表单视图：动态控件生成
│
└── C# 后端服务层（同上）
```

**对比**：

| 维度 | VSTO + WebView2 | VSTO + 纯 WPF |
|------|----------------|---------------|
| UI 开发效率 | **高**（复用前端组件） | 低（全部自研） |
| 看板拖拽 | ✅ 开箱即用 | ❌ 自研 MouseDrag 逻辑 |
| 甘特图 | ✅ 开源组件 | ❌ 自研时间轴 |
| 画册流式布局 | ✅ CSS 原生 | ⚠️ WrapPanel 需虚拟化 |
| 日历 | ✅ FullCalendar | ❌ 自研 |
| 运行时依赖 | WebView2 运行时 | 无额外依赖 |
| 性能（大数据量） | 中（Chromium 开销） | **高**（原生） |
| 开发周期 | 短 | 长 |

> **推荐 VSTO + WebView2**：视图组件不用从零造轮子，UI 体验对标飞书多维表，开发效率最高。

---

## 三、核心模块设计

### 3.1 数据层：Excel ListObject 作为业务数据源

#### 为什么使用 ListObject（超级表）

| 特性 | 普通 Range | ListObject（超级表） |
|------|-----------|---------------------|
| 列名识别 | 需手动管理 | ✅ 自动列名 |
| 行增删 | 复杂 | ✅ ListRows.Add / Remove |
| 数据绑定 | 无 | ✅ 支持数据绑定 |
| 自动扩展 | 无 | ✅ 新增数据自动纳入表范围 |
| 结构化引用 | 无 | ✅ 列名引用 |

**约束**：必须要求用户将数据区域转换为超级表（Ctrl+T），禁止使用普通 Range。

#### 数据读写层设计

```
ExcelAdapter（隔离 Excel Interop）
├── 读取
│   ├── GetData()：批量读取 ListObject → object[,] 数组 → 内存实体集合
│   ├── GetFields()：读取 ListColumns 获取字段名 + 类型推断
│   └── GetRange()：获取表范围用于定位
│
├── 写入
│   ├── UpdateCell(rowIndex, fieldName, value)：单单元格更新
│   ├── AddRow(values)：新增行 ListRows.Add()
│   ├── DeleteRow(rowIndex)：删除行 ListRows.Remove()
│   └── BatchUpdate(updates)：批量更新（合并提交，减少 COM 调用）
│
├── 事件监听
│   ├── Worksheet_Change：监听 Excel 手动修改 → 通知 UI 刷新
│   └── SheetSelectionChange：可选，选中行高亮联动
│
└── 性能控制
    ├── Application.ScreenUpdating = false（批量操作时关闭刷新）
    └── Marshal.ReleaseComObject（及时释放 COM 对象）
```

**关键原则**：
- 全部 Excel COM 操作隔离在 ExcelAdapter 层，上层业务逻辑不直接操作 Excel 对象
- 禁止循环中单单元格读写；批量读取到 `object[,]` 数组，修改后批量回写 `Range`
- 操作完毕后 `Marshal.ReleaseComObject` 释放 COM 对象，避免 Excel 后台残留进程

### 3.2 视图配置存储

#### 方案 A：工作簿自定义 XML 部件（推荐）

VSTO 原生支持将 XML 写入 Excel OpenXML 自定义部件，嵌入 xlsx 文件内部。

```csharp
// 写入自定义 XML 部件
var xmlPart = workbook.CustomXMLParts.Add(configJson);
```

| 优势 | 劣势 |
|------|------|
| ✅ 视图配置跟随文件走，发送 xlsx 即携带配置 | ⚠️ WPS 打开会丢失自定义 XML 部件 |
| ✅ 不污染工作表单元格 | ⚠️ 仅 MS Excel 生效 |
| ✅ 不需要外置文件 | ⚠️ 第三方工具编辑 xlsx 可能丢失 |

#### 方案 B：外置配置文件

同目录生成 `{文件名}.multiview.config.json`。

| 优势 | 劣势 |
|------|------|
| ✅ WPS 也可读取 | ⚠️ 需两个文件一起分发 |
| ✅ 配置可备份 | ⚠️ 复制 xlsx 时容易丢失配置文件 |
| ✅ 不受 Excel 版本影响 | |

#### 视图配置 JSON 结构

```json
{
  "version": "1.0",
  "sourceTable": "Table1",
  "sheetName": "Sheet1",
  "views": [
    {
      "viewId": "v001",
      "viewType": "Kanban",
      "viewName": "任务看板",
      "filter": "[状态] != ''",
      "groupBy": "状态",
      "sort": [{ "field": "截止日期", "order": "asc" }],
      "visibleFields": ["任务名称", "负责人", "截止日期"],
      "cardMeta": {
        "title": "任务名称",
        "image": "附件图片",
        "description": ["负责人", "截止日期"]
      }
    },
    {
      "viewId": "v002",
      "viewType": "Gallery",
      "viewName": "画册浏览",
      "visibleFields": ["任务名称", "图片路径", "负责人"],
      "cardMeta": {
        "title": "任务名称",
        "image": "图片路径",
        "description": ["负责人"]
      }
    },
    {
      "viewId": "v003",
      "viewType": "Calendar",
      "viewName": "日历视图",
      "calendarConfig": {
        "dateField": "截止日期",
        "titleField": "任务名称"
      }
    },
    {
      "viewId": "v004",
      "viewType": "Gantt",
      "viewName": "进度甘特",
      "ganttConfig": {
        "startField": "开始日期",
        "endField": "截止日期",
        "labelField": "任务名称"
      }
    }
  ]
}
```

### 3.3 双向数据同步机制

这是 VSTO 插件方案的核心难点，需要处理 Excel ↔ UI 双向同步：

```
┌──────────────────────────────────────────────────────────┐
│                    双向同步流程                           │
│                                                          │
│  Excel 手动修改单元格                                     │
│       ↓                                                   │
│  Worksheet_Change 事件触发                                │
│       ↓                                                   │
│  EventBridge 接收 → 防抖处理（避免循环触发）              │
│       ↓                                                   │
│  ExcelAdapter 重新读取变更数据                            │
│       ↓                                                   │
│  通过 WebView2 PostMessage 推送给前端                     │
│       ↓                                                   │
│  Vue 前端视图刷新                                          │
│                                                          │
│  ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─  │
│                                                          │
│  用户在 UI 视图编辑（看板拖拽 / 表单输入 / 甘特拖拽）       │
│       ↓                                                   │
│  Vue 前端通过 WebView2 发送编辑指令                       │
│       ↓                                                   │
│  C# Bridge 接收 → DataSyncManager 处理                    │
│       ↓                                                   │
│  ExcelAdapter.UpdateCell 写入 ListObject 单元格          │
│       ↓                                                   │
│  触发 Excel 原生变更 → 其他视图自动刷新                   │
│                                                          │
│  ⚠️ 防抖关键：                                            │
│  UI 编辑写入 Excel → 触发 Change 事件 → 通知 UI 刷新      │
│  必须加标志位防止循环：isUpdatingFromUI                    │
└──────────────────────────────────────────────────────────┘
```

**事件防抖设计**：

```csharp
private bool _isUpdatingFromUI = false;

private void Worksheet_Change(Range target)
{
    if (_isUpdatingFromUI) return;  // UI 触发的变更，不再回传 UI
    
    // 防抖：延迟 300ms 合并多次变更
    _debounceTimer.Change(300, Timeout.Infinite);
}

public void UpdateCellFromUI(int rowIndex, string field, object value)
{
    _isUpdatingFromUI = true;
    try
    {
        _excelAdapter.UpdateCell(rowIndex, field, value);
    }
    finally
    {
        _isUpdatingFromUI = false;
    }
}
```

### 3.4 UI 视图层实现

#### WebView2 桥接通信

```csharp
// C# → JS 调用
webView.CoreWebView2.PostWebMessageAsJson(jsonMessage);

// JS → C# 接收
webView.CoreWebView2.WebMessageReceived += (s, e) => {
    var message = JsonSerializer.Deserialize<BridgeMessage>(e.WebMessageAsJson);
    HandleBridgeMessage(message);
};
```

```javascript
// Vue 前端侧
// 接收 C# 消息
window.chrome.webview.addEventListener('message', (event) => {
  const message = JSON.parse(event.data);
  handleBridgeMessage(message);
});

// 发送消息给 C#
function sendToHost(message) {
  window.chrome.webview.postMessage(JSON.stringify(message));
}

// 调用示例：获取数据
sendToHost({ action: 'getData', viewId: 'v001' });

// 调用示例：更新单元格
sendToHost({ 
  action: 'updateCell', 
  rowIndex: 5, 
  field: '状态', 
  value: '已完成' 
});
```

#### 六大视图在 TaskPane 中的实现

| 视图 | 前端组件 | 与 Excel 交互 |
|------|---------|-------------|
| 表格视图 | vxe-table / AG Grid | 双击编辑 → 回写 ListObject 单元格 |
| 表单视图 | 动态表单组件 | 提交 → 批量更新行 → Excel |
| 看板视图 | vue-draggable | 拖拽跨列 → 修改分组字段 → Excel |
| 画册视图 | CSS Grid + 图片懒加载 | 图片路径字段 → 加载本地图片 |
| 日历视图 | FullCalendar | 拖拽事件 → 修改日期字段 → Excel |
| 甘特视图 | frappe-gantt | 拖拽进度条 → 修改起止日期 → Excel |

### 3.5 架构分层

```
VSTO Excel Add-in
├── ExcelAdapter 层 【隔离 Excel Interop】
│   └── 读取 ListObject、批量读写单元格、自定义 XML 部件读写
│   └── 监听工作表事件（Change、SelectionChange）
│
├── Core 业务层
│   ├── DataModel：行实体；FieldSchema：字段元信息
│   ├── ViewConfig：视图配置模型
│   └── ViewEngine：原始数据 + 视图配置 → 过滤/分组/排序后的视图数据集
│
├── UIBridge 层
│   └── WebView2 C# ↔ JS 双向消息通信
│   └── 事件防抖、循环触发抑制
│
└── UI 层（Vue3 前端 / WPF UserControl）
    └── 表格、表单、看板、画册、日历、甘特视图渲染与交互
```

**核心设计原则**：把 Excel Interop 全部隔离在 ExcelAdapter 层，上层业务层不直接操作 Excel COM 对象，便于测试和维护。

---

## 四、开发难度评估

### 4.1 模块难度分解

| 模块 | 难度 | 说明 | 你的经验匹配度 |
|------|------|------|---------------|
| Excel 数据交互层 | **低-中** | VSTO API 成熟，ListObject 读写资料多 | ✅ 高（VBE2021 经验） |
| 视图配置序列化持久化 | **低** | JSON 序列化 + OpenXML 自定义部件 | ✅ 高 |
| WebView2 嵌入 + C#/JS 桥接 | **中** | 双向通信、消息序列化、异步处理 | ⚠️ 中（需学习 WebView2 API） |
| 事件同步：Excel ↔ UI 双向刷新 | **中** | 事件防抖、循环触发抑制 | ⚠️ 中 |
| 表格视图 UI | **低** | 前端表格组件开箱即用 | — |
| 表单视图 UI | **低** | 动态控件生成 | — |
| 看板视图 UI | **中** | 前端拖拽组件，需对接回写 Excel | — |
| 画册视图 UI | **中** | 流式布局 + 虚拟化 | — |
| 日历视图 UI | **中** | FullCalendar 集成 + 日期回写 | — |
| 甘特视图 UI | **中-高** | 时间轴拖拽 + 日期回写 Excel | — |

### 4.2 总体难度评级

| 维度 | 评级 | 说明 |
|------|------|------|
| 数据层 | ⭐⭐☆☆☆ | Excel ListObject API 成熟，你有 VSTO 经验 |
| 通信层 | ⭐⭐⭐☆☆ | WebView2 桥接需要处理双向通信和防抖 |
| UI 层（WebView2 路线） | ⭐⭐⭐☆☆ | 复用前端组件，不用从零写 |
| UI 层（纯 WPF 路线） | ⭐⭐⭐⭐⭐ | 看板拖拽、甘特时间轴全部自研 |
| 部署分发 | ⭐⭐⭐⭐☆ | VSTO ClickOnce 部署门槛高 |

### 4.3 与独立 EXE 方案难度对比

| 对比维度 | 独立 EXE | Excel 插件 |
|---------|---------|-----------|
| 文件解析层 | ⭐⭐⭐⭐（自建 xlsx 解析） | ⭐⭐（复用 Excel 引擎） |
| 数据存储 | ⭐⭐⭐（自建文件 IO） | ⭐（直接用 Excel） |
| UI 视图层 | ⭐⭐⭐（复用前端组件） | ⭐⭐⭐（同） |
| 双向同步 | ⭐⭐（内存数据集，简单） | ⭐⭐⭐⭐（Excel ↔ UI 事件同步复杂） |
| 部署分发 | ⭐⭐（单 exe） | ⭐⭐⭐⭐（VSTO 部署复杂） |

---

## 五、开发风险与对策

| 风险 | 等级 | 影响 | 对策 |
|------|------|------|------|
| **VSTO 分发部署门槛高** | 🔴 高 | ClickOnce 安装失败、需要 VSTO 运行时、Office 版本匹配 | 使用 Windows Installer (MSI) 替代 ClickOnce；提供详细安装文档 |
| **Office 版本差异** | 🟡 中 | 32/64 位 Office 的 Interop 行为差异 | 明确支持 Office 版本范围；32/64 位分别测试 |
| **WebView2 运行时缺失** | 🟡 中 | 目标机器无 Edge/WebView2 | 打包时捆绑 WebView2 固定版本引导安装 |
| **自定义 XML 部件丢失** | 🟡 中 | WPS 打开 xlsx 丢失视图配置 | 检测 WPS 环境并提示用户；备选外置 JSON |
| **COM 对象残留** | 🟡 中 | Excel 后台进程残留 | 严格 `Marshal.ReleaseComObject`；使用 using 模式 |
| **事件循环触发** | 🟡 中 | UI 编辑 → Excel Change → UI 刷新 → 循环 | `_isUpdatingFromUI` 标志位 + 防抖定时器 |
| **WPS 不兼容** | 🔴 高 | VSTO 插件不能在 WPS 运行 | 明确标注仅支持 MS Excel；推荐独立 EXE 方案覆盖 WPS 用户 |
| **TaskPane 尺寸限制** | 🟢 低 | 侧边栏宽度有限，视图展示空间受限 | 支持可拖拽浮出窗口；全屏模式 |
| **大数据量性能** | 🟡 中 | 连续窗体重绘、大量 COM 调用卡顿 | 批量读取到数组；Application.ScreenUpdating 控制；虚拟列表 |

---

## 六、实现路线：两种 UI 形态

### 6.1 形态一：侧边栏 TaskPane（标准 VSTO 方案）

```
┌─────────────────────────────────────────────┐
│  Excel 主窗口                                │
│  ┌─────────────────────────┬──────────────┐ │
│  │                         │  TaskPane    │ │
│  │   Excel 工作表           │  (多维表     │ │
│  │   (超级表 ListObject)    │   视图区)    │ │
│  │                         │              │ │
│  │                         │  ┌──────────┐│ │
│  │                         │  │ 看板/画册││ │
│  │                         │  │ /日历/甘特││ │
│  │                         │  │  视图     ││ │
│  │                         │  └──────────┘│ │
│  │                         │  视图切换标签 │ │
│  └─────────────────────────┴──────────────┘ │
└─────────────────────────────────────────────┘
```

**特点**：
- TaskPane 停靠在 Excel 侧边，宽度可调整
- Excel 工作表和多维表视图同时可见
- 在 Excel 中修改数据 → 侧边视图实时刷新
- 在侧边视图中拖拽卡片 → Excel 单元格更新

**优势**：用户不需要离开 Excel，数据源和视图同屏可见

**劣势**：侧边栏宽度有限，看板多列、画册网格展示空间受限

### 6.2 形态二：独立浮出窗口（自定义界面）

VSTO 支持创建独立 WPF 窗口，不限于 TaskPane 侧边栏：

```
┌─────────────────────────────────────────────┐
│  Excel 主窗口（最小化/后台）                   │
│  ┌─────────────────────────┐                │
│  │  Excel 工作表             │                │
│  │  (超级表 ListObject)      │   ┌──────────────────────────┐
│  │                          │   │  多维表独立窗口 (WPF)       │
│  │                          │   │  ┌──────────────────────┐ │
│  │                          │   │  │   WebView2 / WPF UI  │ │
│  │                          │   │  │                      │ │
│  │                          │   │  │  完整 6 大视图展示    │ │
│  │                          │   │  │  无侧边栏尺寸限制      │ │
│  │                          │   │  └──────────────────────┘ │
│  │                          │   └──────────────────────────┘
│  └─────────────────────────┘
└─────────────────────────────────────────────┘
```

**实现方式**：
- VSTO Add-in 中创建独立 `Window`（WPF）或 `Form`（WinForms）
- 窗口内嵌入 WebView2 或 WPF UserControl
- 通过 ExcelAdapter 读写后台 Excel 工作表数据
- 窗口可全屏、可调整大小，不受 TaskPane 限制

**优势**：
- 界面自由度最高，等同于"在 Excel 进程内运行的独立窗口"
- 可以全屏展示看板、画册、甘特图
- 用户体验接近独立软件，但数据源是 Excel

**劣势**：
- 需要额外管理窗口生命周期（Excel 关闭时同步关闭）
- 用户可能混淆"这个窗口属于 Excel 还是独立程序"

### 6.3 两种形态对比

| 维度 | TaskPane 侧边栏 | 独立浮出窗口 |
|------|----------------|-------------|
| 展示空间 | 受限（侧边宽度） | **充足**（全屏可用） |
| Excel 数据同屏可见 | ✅ 是 | ❌ 否（需切换窗口） |
| 开发复杂度 | 低（标准 VSTO） | 中（管理窗口生命周期） |
| 用户体验 | 数据源+视图同屏 | 接近独立软件 |
| 适合视图 | 表格、表单、简单看板 | **全部 6 大视图** |

> **推荐**：默认使用独立浮出窗口形态（形态二），获得最大的 UI 展示空间；可选提供 TaskPane 模式作为紧凑布局。

---

## 七、与独立 EXE 方案对比总结

### 7.1 核心对比

| 对比维度 | 独立 EXE | Excel 插件（VSTO + WebView2） |
|---------|---------|-------------------------------|
| **宿主依赖** | 无，完全独立 | 绑定 MS Excel |
| **数据存储** | 自建文件解析（EPPlus/NPOI） | 复用 Excel ListObject |
| **文件兼容** | xlsx/csv 通用 | 仅 Excel 超级表 |
| **WPS 兼容** | ✅ 不涉及 | ❌ 完全不兼容 |
| **分发部署** | 单 exe，无依赖 | ClickOnce/MSI，需 Office 运行时 |
| **开发效率** | 中（需自建文件 IO） | 中高（复用 Excel 数据引擎） |
| **双向同步复杂度** | 低（内存数据集） | 高（Excel COM 事件 ↔ UI） |
| **UI 自由度** | 完全自定义 | TaskPane 受限 / 独立窗口自由 |
| **用户迁移成本** | 需学习新软件 | 零迁移（用户已在 Excel） |
| **Excel 原生功能共存** | ❌ 不可共存 | ✅ 公式、条件格式、图表共存 |
| **exe 体积** | 30-150 MB | 10-30 MB（不含 Office） |
| **适合场景** | 外部分发、WPS 用户、完全离线 | 内部团队、深度 Excel 用户 |

### 7.2 决策矩阵

| 你的情况 | 推荐方案 |
|---------|---------|
| 需要分发给外部客户，Office 版本不可控 | **独立 EXE** |
| 需要 WPS 兼容 | **独立 EXE** |
| 内部团队，统一 MS Office，深度依赖 Excel | **Excel 插件** |
| 需要 Excel 公式/条件格式与视图共存 | **Excel 插件** |
| 追求最简单分发 | **独立 EXE** |
| 追求最快开发出可用产品 | **Excel 插件**（复用 Excel 数据引擎） |
| 两者都要覆盖 | **先做独立 EXE，后做 Excel 插件**（共享 ViewEngine 核心） |

### 7.3 混合策略建议

如果两种方案都想覆盖，建议采用**共享核心层**的架构：

```
共享核心层（.NET 类库）
├── ViewEngine（视图引擎：过滤、排序、分组）
├── ViewConfig 模型（视图配置数据结构）
├── DataModel（内存数据模型）
└── 通用工具类

        ↗ 独立 EXE 适配层
       /    └── FileAdapter（EPPlus/NPOI 读写 xlsx）
      /
核心层
      \
       ↘ Excel 插件适配层
           └── ExcelAdapter（ListObject 读写 + 事件监听）
```

**优势**：
- ViewEngine 核心逻辑只写一次，两种方案共享
- 数据模型统一，降低维护成本
- 优先开发独立 EXE（覆盖面广），再追加 Excel 插件（利用已有核心层）

---

## 八、开发实施顺序（Excel 插件方案）

### 阶段 1：基础底座

| 任务 | 说明 |
|------|------|
| VSTO 项目搭建 | .NET 8 VSTO Excel Add-in，Custom TaskPane 创建 |
| ExcelAdapter | ListObject 读取、批量读写、自定义 XML 部件读写 |
| WebView2 嵌入 | TaskPane 内放置 WebView2，加载 Vue 前端 |
| C#/JS 桥接 | PostMessage 双向通信通道搭建 |

### 阶段 2：核心通路

| 任务 | 说明 |
|------|------|
| ViewEngine | 过滤、排序、分组逻辑（可复用独立 EXE 方案的核心层） |
| 表格视图 | 前端表格组件展示 Excel 数据，行内编辑回写 |
| 双向同步 | Worksheet_Change → UI 刷新；UI 编辑 → Excel 回写 |
| 事件防抖 | 循环触发抑制、防抖定时器 |

### 阶段 3：交互视图

| 任务 | 说明 |
|------|------|
| 看板视图 | 分组 + 拖拽 + 回写分组字段 |
| 画册视图 | 流式卡片 + 图片加载 |
| 表单视图 | 动态控件 + 批量更新 |
| 视图管理 | 新建/编辑/删除视图，配置持久化 |

### 阶段 4：时间维度

| 任务 | 说明 |
|------|------|
| 日历视图 | FullCalendar + 日期回写 |
| 甘特视图 | 进度条 + 拖拽改日期回写 |

### 阶段 5：完善与发布

| 任务 | 说明 |
|------|------|
| 独立窗口模式 | 浮出窗口实现（不受 TaskPane 限制） |
| 性能优化 | 批量操作、ScreenUpdating 控制、COM 释放 |
| 安装包 | MSI 安装包 + VSTO 运行时引导 |
| 兼容测试 | Office 2016/2019/365 × 32/64 位 |

---

## 九、VSTO 开发避坑要点

| 要点 | 说明 |
|------|------|
| **COM 对象释放** | 所有 Excel COM 对象操作后及时 `Marshal.ReleaseComObject`，避免 Excel 后台残留进程 |
| **事件防抖** | Worksheet_Change 频繁触发，加防抖 + `_isUpdatingFromUI` 标志位防止循环刷新 |
| **批量操作** | 不在循环中单单元格读写；全部读取到 `object[,]` 数组，修改后批量写回 `Range` |
| **ScreenUpdating** | 批量操作前 `Application.ScreenUpdating = false`，操作后恢复 |
| **视图配置与数据分离** | 视图层只做过滤展示，永远不修改原始 Excel 行顺序、原始筛选 |
| **图片存储** | 不嵌入图片到单元格，只存储本地文件路径，UI 视图负责加载图片 |
| **线程安全** | UI 线程与 Excel COM 线程不同，跨线程操作 Excel 需 Dispatcher |
| **WPS 兼容性** | 明确标注仅支持 MS Excel，自定义 XML 部件在 WPS 中会丢失 |

---

## 十、总结

| 项目 | 决策 |
|------|------|
| 可行性 | ✅ 技术可行 |
| 推荐技术组合 | VSTO (.NET 8) + WebView2 + Vue3 前端 |
| 数据源 | Excel ListObject（超级表） |
| 视图配置存储 | 自定义 XML 部件（首选）/ 外置 JSON（备选） |
| UI 形态 | 独立浮出窗口（推荐）/ TaskPane 侧边栏 |
| 核心优势 | 复用 Excel 数据引擎、零用户迁移成本、与 Excel 原生功能共存 |
| 核心风险 | 绑定 MS Excel、WPS 不兼容、VSTO 部署复杂、双向同步事件防抖 |
| 与独立 EXE 关系 | 可共享 ViewEngine 核心层，两种方案互补覆盖不同用户群 |

**最终建议**：

1. **如果目标用户主要使用 MS Excel** → Excel 插件方案可行且开发效率高，你的 VSTO 经验直接可用
2. **如果需要覆盖 WPS 用户或外部客户** → 独立 EXE 方案更合适
3. **最优策略** → 先开发独立 EXE（覆盖面广），构建共享核心层（ViewEngine + 数据模型），再追加 Excel 插件版本（复用核心层 + ExcelAdapter 适配层），两个产品共享核心逻辑，分头适配不同宿主

---

*本文档基于需求分析与技术可行性调研生成，可与《开发规划文档与技术路线.md》对照阅读。*
