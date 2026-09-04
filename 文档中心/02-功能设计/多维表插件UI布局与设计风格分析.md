# 多维表分析插件 — 界面 UI 布局与设计风格分析

> 对象：`MultiTableAddin`（Excel-DNA + .NET 8 + WebView2 单文件 SPA 插件）
> 分析依据：`MultiTableAddin/src/MultiTableAddin/Html/index.html`（真实 UI 全部在此内嵌）
> 版本基线：v2.9.42（2026-08-20）

---

## 一、整体架构与渲染载体

| 层 | 技术 | 职责 |
|---|---|---|
| 宿主 | C# WinForms `HtmlMainForm` | 承载 WebView2 控件、窗口标题/图标/置顶/跟随 Excel 最小化、JS↔C# 桥接 |
| 渲染 | WebView2 + Chromium | 通过 `NavigateToString(LoadEmbeddedHtml())` 加载**内嵌资源** `MultiTableAddin.Html.index`（每次从 DLL 读取，无磁盘缓存陈旧问题） |
| 逻辑 | 单文件 HTML + 内联 CSS + 原生 JS（无框架） | 全部界面、视图、交互、配置读写 |
| 数据源 | Excel/WPS 超级表（ListObject） | 经 C# `ExcelAdapter` 读取行列，JS 通过 `api.call(...)` 调用 |

**关键结论**：真实可见的 UI 100% 在 `index.html`。C# 侧的 `MultiTableMainView.xaml`/`KanbanView.xaml`/`AddViewDialog` 等为历史死代码（BAML 资源仍编译进 DLL，但不影响 WebView2 渲染）。任何"改界面"只需改 `index.html` 的 HTML/CSS/JS。

---

## 二、页面骨架（DOM 布局）

```
#app
└─ .body  (横向 flex：左侧栏 + 右侧主区)
   ├─ .sidebar  #sidebar  (宽 115px，固定左列)
   │  ├─ .sidebar-top  #sidebarTop   (设置/刷新/帮助 三个全宽按钮)
   │  └─ #sidebarViews                (视图列表 .view-item)
   └─ .main  (纵向 flex)
      ├─ .viewbar  #viewbar           (高 32px 顶部工具条：视图标签 .vtabs + 工具 .tools)
      └─ .stage  (纵向 flex：内容 + 右栏)
         ├─ .content  #content        (各视图渲染区)
         └─ .right-col  #rightCol     (搜索结果 / 详情面板 / 设置面板 三选一)
```

悬浮层（全局，脱离 .body 流）：
- `#fieldPanel` 字段设置面板
- `#ctxMenu` 右键上下文菜单
- `.overlay/.modal` 模态确认/输入框
- `#toast` 轻提示

**布局特征**：经典"左窄边栏 + 右主区"三栏式（侧栏 / 内容 / 右栏），参考飞书多维表。侧栏固定 115px 仅容纳图标+视图名；主区顶部 32px 工具条；内容区与右栏可联动（选中行→右栏详情）。

---

## 三、设计语言（Design Tokens）

全部以 CSS 自定义属性（`:root` 变量）统一管理，便于换肤与一致性：

### 3.1 色彩系统

| 变量 | 值 | 用途 |
|---|---|---|
| `--accent` | `#3370FF` | 主蓝：激活态、视图图标、强调按钮、选中行、表头下划线 |
| `--accent-dark` | `#173A8C` | 深蓝变量（v2.9.42 起所有小图标统一改用 `--accent`，此变量仅保留作标题 / 大块底色） |
| `--accent-weak` | `#E8F0FE` | 浅蓝底：选中行/激活标签背景 |
| `--accent-weak2` | `#E1EAFF` | 更浅蓝底：悬停高亮 |
| `--text` / `--text-2` / `--text-3` | `#1D2129` / `#4E5969` / `#86909C` | 三级文字灰阶 |
| `--bg` / `--bg-2` / `--bg-3` | `#FFFFFF` / `#F7F8FA` / `#F2F3F5` | 三级背景 |
| `--border` / `--border-strong` | `#E5E6EB` / `#C9CDD4` | 边框 |
| `--hover` | `#F2F3F5` | 悬停底 |
| `--danger` / `--ok` | `#F53F3F` / `#00B42A` | 危险/成功 |
| `--bar` | `#3370FF,#14C9C9,#FF7D00,#F76965,#9FDB60,#CA62E8,#FFC53D,#6E4AE0` | 画册分组 8 色泳道调色板 |

**配色哲学**：低饱和度商务蓝为主轴（#3370FF），辅以中性灰阶；强调色仅用于交互态与数据可视化，避免视觉噪音。整体偏"克制、信息密度高"的办公风。

### 3.2 圆角与尺寸

- `--radius: 6px`：全局圆角（按钮、输入框、卡片、菜单）。
- 按钮 `.btn` 高 26px、`.btn.sm` 高 24px；侧栏宽 115px；工具条高 32px；图标 14×14px。
- 字号：基准 13px，按钮/工具条 12px，分组标题 12px，视图名 14px/500。

### 3.3 排版

- 系统字体栈，中文优先；行高紧凑（按钮 `line-height:1`）。
- 文本溢出统一 `text-overflow:ellipsis; white-space:nowrap`（视图名、卡片描述、汇总值）。
- 编号/序号列 `.col-rownum` 居中、灰底。

---

## 四、视图体系（核心能力）

`ViewType` 枚举：Table / Form / Gallery / Calendar / Gantt / Dashboard / Chart（**Kanban 视图类型已于 v2.9.42 移除**，遗留配置打开时自动迁移为 Gallery）。

### 4.1 视图清单与渲染入口

| 视图 | 渲染函数 | 关键样式类 | 说明 |
|---|---|---|---|
| 表格 | `renderTable` | `table.grid` / `thead th` sticky | 主数据网格，支持汇总脚 `tfoot`、列宽拖拽 `.resizer`、排序 `.sort.asc/desc`、行内编辑 `.cell-edit` |
| 表单 | `renderForm` | `.form-row` / `.flabel` / `.fval` | 单记录分页浏览，左标签右值，支持上一条/下一条、查找过滤 |
| 画册 | `renderGallery` | `.kanban-cols` / `.kanban-col` / `.gallery-card` | 卡片瀑布；**分组模式复用分组列布局与拖拽机制**（`enableKanbanDrag`，单击仅选中不弹详情） |
| 日历 | `renderCalendar` | `.cal-event-seg`（绝对定位浮层段） | 月视图重叠事件泳道分配（最大团着色），动态行高，+N 点击展开该周 |
| 甘特 | `renderGantt` | `.gantt-left` / `.gantt-ctrl-row` / `.gantt-overlay` | 区间条着色、多列控制行、today 线、上层标签 |
| 仪表盘 | `renderDashboard` | `.stat` 卡 + LiveCharts2 | 指标卡 + 图表网格，`dashTheme` 切换图表配色（office 等多套） |
| 图表 | （单视图统计） | LiveCharts2 | 不在侧栏列表显示，`renderSidebar` 过滤 `viewType==='Chart'` |

### 4.2 视图元数据

- `VIEW_ICON`：视图→图标名映射（Gallery 复用 `kanban` 图标，v2.9.40）。
- `FIELD_ICON`：字段→emoji 映射（▦▤▥▨▣）。
- `VIEW_LABEL`：视图类型→中文名。
- `ICONS`：内联 SVG 图标库（table/refresh/pin/info/form/kanban/gear/calendar/gantt/dashboard/chart/help/copy），统一 `stroke="currentColor"`，由 CSS `color` 着色。

### 4.3 视图列表交互

- 自带 7 类基础视图不可删/改名；用户视图（`userView=true`，持久化于 C# `ViewConfig`）可删/改名。
- 右键菜单：另存新视图 / 重命名 / 删除（`#ctxMenu` 复用容器，仿记录右键菜单）。
- 首个用户视图前插 `.view-sep` 分隔线区分自带/用户。
- 悬停 `it.title` 显示完整视图名（侧栏窄会截断）。

---

## 五、组件库（复用单元）

| 组件 | 类名 | 要点 |
|---|---|---|
| 按钮 | `.btn` / `.btn.sm` / `.btn.primary` / `.btn.ghost` / `.btn.active` | 边框+底，primary 为蓝底白字 |
| 图标位 | `.ic` | 14×14 内联 SVG，`currentColor` 着色 |
| 视图标签 | `.vtab` / `.vtab.active` | 工具条内圆角标签 |
| 视图项 | `.view-item` / `.view-item.user` / `.view-item.active` | 自带蓝、用户灰、激活蓝 |
| 画册分组列 | `.kanban-col` / `.kanban-col.dropover` | 画册分组列（看板式列布局），拖拽落点高亮虚线框 |
| 卡片 | `.kanban-card` / `.gallery-card` / `.dragging` / `.selected` | 悬停/拖拽/选中态 |
| 表单行 | `.form-row` / `.flabel` / `.fval` | 标签 100px 固定 + 值 |
| 网格 | `table.grid` / `thead th` sticky / `tfoot` 汇总 | 冻结表头、汇总脚 |
| 模态 | `.overlay` / `.modal` / `.modal-head` / `.modal-body` | 遮罩+居中弹窗，含 ✕ 关闭 |
| 右键菜单 | `#ctxMenu` / `.ctx-menu` | 定位弹层 |
| 轻提示 | `#toast` | 1.8s 自动隐藏 |
| 字段面板 | `#fieldPanel` / `.field-panel` | 字段设置悬浮面板 |

---

## 六、交互模式

1. **桥接通信**：JS `api.call(method, params)` → C# `WebView2.WebMessageReceived` 派发 → 返回 Promise。所有数据读写经此通道。
2. **拖拽**：`enableKanbanDrag`（mousedown/move/up + `document.elementsFromPoint` + `closest('.kanban-col')`）；v2.9.40 起拖到末尾空白处归入最后一列；画册分组复用同一套。
3. **右键菜单**：记录右键、视图项右键均复用 `#ctxMenu`，仿飞书定位。
4. **模态确认**：`.overlay` 模态（重命名/删除/确认），ESC 取消关闭。
5. **轻提示**：`toast(msg)` 统一反馈。
6. **字段自动识别**：按字段类型（Select/Quarter/Text/Image/DateTime/Number...）匹配视图默认配置；`ensureViewDefaults` 补全缺失字段。
7. **配置持久化**：一簿一文件 `{filename}.multiview.json` + 嵌入隐藏表 `_MultiTableConfig`（VeryHidden）；`saveLocation`=excel/file/both。

---

## 七、值得沿用/复用的设计资产（供桌面版继承）

- **CSS Token 体系**：`:root` 变量已成熟，可直接迁移作为桌面版设计系统底座。
- **视图框架**：`ViewConfig`/`ViewType`/`ensureViewDefaults` 的"多视图绑定数据源"模型与渲染无关，可解耦数据层后复用。
- **组件库**：`.btn/.view-item/.kanban-col/.modal/.toast/.ctx-menu` 等已成型，桌面版前端可整体继承。
- **拖拽/右键/模态/轻提示**交互范式成熟，可作为桌面版交互基线。
- **图标体系**：内联 SVG + `currentColor` 方案，无外部依赖、易换色。

## 八、当前局限（桌面版需改进点）

1. 强依赖 Excel/ListObject 作为数据源——无独立存储，脱离 Excel 即不可用。
2. 配置嵌入工作表，跨设备/跨文件迁移不便。
3. 无 Excel 导入导出（数据进出全靠 Excel 本身）。
4. 单进程单文件，无多库/多项目组织能力。
（以上正是"桌面独立软件"方案的改造方向，详见《桌面版独立软件设计方案》。）
