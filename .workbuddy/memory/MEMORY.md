# 多维表分析软件 — 项目记忆

## 技术路线
- Excel-DNA 1.9(patched) + .NET 8 (net8.0-windows) + WPF + HandyControl 3.5.1 + WebView2；视图=表格/表单/看板/画册/日历/甘特/仪表盘/图表；数据源=Excel/WPS 超级表(ListObject)。模型已支持多视图(`ViewConfigFile.Views`/`List<ViewConfig>`)，无需改 C# 架构。
- **真实 UI 是 WebView2(`HtmlMainForm`) 加载内嵌 `index.html`**；WPF `MultiTableMainView`/`ChartView`/`CalendarView` 等是死代码，**改可见控件只能改 index.html 的 HTML/JS**。

## 构建与交付（关键易错）
- `dotnet` 在 `C:/Program Files/dotnet/dotnet.exe`。
- `dotnet build -c Release` 产物在 `MultiTableAddin/src/MultiTableAddin/bin/Release/net8.0-windows/`；**安装从 `MultiTableAddin/dist/files/` 取 dll**。**只 build 不更新 dist**。交付：先 build，再 `rm -f MultiTableAddin/dist/files/MultiTableAddin.dll && cp .../bin/Release/.../MultiTableAddin.dll MultiTableAddin/dist/files/`（dist 旧 dll 常句柄锁，先 rm 再 cp）。
- `tools/build_addin.ps1` 因中文路径在 PowerShell `&` 下路径乱码无法运行 → 改用 Git Bash 直接 dotnet build + cp。
- **交付前必跑 `python verify-build.py`**（exit 0 才算过；版本号+检查串+负向守卫）。

## verify-build.py 检查串铁律（三分法）
- ① C# 字符串字面量(const/版本号)→ #US 堆 = **UTF-16LE**；② C# 类型/方法/属性名 → 元数据 #Strings 堆 = **UTF-8**(写成 utf-16le 会误报 FAIL)；③ 内嵌 HTML/JS 字符串 → **UTF-8**。注释不进 DLL。
- 改 index.html 后**必须同步 verify-build.py 检查串**（失效项修正 + 新增本回合功能串 + 必要时加 negative_checks 负向断言→已改为"旧代码已移除"）。功能需 build+copy 才生效。
- **index.html 现为 LF（上轮 re.sub 编辑由 CRLF 归一化）**：verify-build 跨行检查串必须写 `\n`，写 `\r\n` 会误报 FAIL。
- **裸 `kanban` 不能作负向守卫**：死代码 `KanbanView.xaml`(BAML 资源) 恒成立 "kanban"，且 JS `VIEW_ICON` 保留 `Kanban:'kanban'`/`Gallery:'kanban'` → 裸 `kanban`(utf-16le/utf-8) 永不缺失。看板删除以"renderKanban 函数已删"负向 + "Gallery:'kanban'"正向覆盖。

## WebView2 铁律
- 本地图片：`<img src="file://">` 被拦，须经 C# 读文件返 base64 data URI。
- 整页截图：用 CDP `Page.captureScreenshot`(captureBeyondViewport:true, clip 整页坐标)，截图前 JS 解锁滚动容器 `height:auto;overflow:visible`。`CapturePreviewAsync` 只截视口。
- 原生 `<input type=date>` 弹层样式/滚轮不可改，用自建 `attachDatePicker`。
- 自定义弹层事件监听只在 `show()` 注册一次，勿在 `render()` 重复 addEventListener。

## 配置系统（v2.9.35+）
- 一簿一文件 `{filename}.multiview.json`；根 `WorkbookConfigFile{LastTableName,Tables,Order}`。
- 嵌入隐藏工作表 `_MultiTableConfig`(xlSheetVeryHidden=2，用户不可见)，`WriteConfigSheet` 按 30000 字符/块分块写 A 列；读优先级 隐藏表→外部JSON。
- 保存位置开关 `saveLocation`=excel/file/both(默认)；`file` 会删嵌入表防读陈旧。
- VeryHidden 表用户不可达 → 软件内置「查看配置结构与数据」入口(`getEmbeddedConfig`/`setConfigSheetVisible`/`saveWorkbookFile`)。
- `SourceFile` 只存文件名(`Path.GetFileName`)，不写绝对路径(#467)。
- 兼容：`ViewConfigManager.Load/Save` 签名不变只改内部(30+ 死代码调用点)。

## 月视图布局铁律（v2.9.35+）
- 事件绝对定位悬浮段 `.cal-event-seg`（非单元格内嵌）；`layoutMonthEvents` 用 `offsetLeft/offsetTop` 在 rAF 后定位。
- 行高 `gridTemplateRows=表头30px + 各周(展开周显式px/其余1fr)`；+N 点击=展开该周(非弹列表)。
- 动态泳道=区间图着色(最大团)，无写死 5 上限；网格 `align-content:start;min-height:0`；周末列 `#F7F7F7`(getDay 0/6)。
- 画册分组复用看板 `kanban-cols`+`kanban-col`；必须 `renderGalleryInner` 读 `viewGroupField(v)` 真落地。

## 其它铁律
- CSS 类名冲突最隐蔽(#421)：新控件命名完全独立、不共享基类；握把用 `::before` 双竖纹，勿用 `⟨⟩`/`«»` 等非常用 Unicode(变方块)；拖拽吸附=网格单元(`Math.round(dx/cellW)`)。
- flex 列内卡片高度不齐真凶是 `flex-shrink:1`(#465)：子项显式 `flex:0 0 auto`。

## 图标着色与看板下线铁律（v2.9.41）
- **"深蓝"≠越深越好**：`--accent-dark:#173A8C` 在 14px 小图标下肉眼近似黑，用户会报"图标是黑色"。需肉眼可辨蓝的图标一律用 `--accent:#3370FF`（与 `.view-item .ic` 视图列表图标一致），`--accent-dark` 留给标题/大块。**不要给侧栏顶部功能区按钮加品牌文案块**（`.sidebar-brand`）：用户明确不想要"多维表分析"文字标签，左上角只要彩色图标即可。
- **删视图类型≠reroute 渲染**：删 `renderKanban` + `renderContent` reroute 只改渲染，不改侧栏列表——持久化配置里的 Kanban 视图仍会被 `renderSidebar` 的 `forEach` 列出（它只过滤 `Chart`）。彻底下线须在 `ensureViewDefaults` 的 `forEach` 开头迁移：`if(v.viewType==='Kanban'){ v.viewType='Gallery'; if(v.viewName==='看板视图'||!v.viewName) v.viewName='画册视图'; changed=true; }`，靠 `changed` 触发 `saveConfig()` 持久化（下次开不再有该类型）。
- **HTML 内嵌机制**：`HtmlMainForm` 经 `NavigateToString(LoadEmbeddedHtml())` 从内嵌资源 `MultiTableAddin.Html.index` 每次新读（无磁盘缓存陈旧问题）；C# 窗口标题 `Text="多维表分析"`（用户口中"窗口标题"）。判 v2.9.x 是否真部署：看用户能否看到本轮新增的可见 DOM（如 v2.9.40 的 `.sidebar-brand` 文案）。

## 视图侧栏右键菜单（v2.9.38）
- 侧栏每个视图项右键→菜单：另存新视图(所有视图)/重命名(仅用户视图)/删除(仅用户视图)。自带 8 类基础视图不可删、不可改名。
- **`UserView` 标志存于 C# `ViewConfig`(持久化字段，旧配置缺省=false=基础视图，向后兼容)**；勿用 JS 临时字段标记——C# 强类型反序列化会丢弃，重开文件即丢失标记。
- `duplicateView` 设 `copy.userView=true`；JS 甘特自动补建也 `userView:true`。`isUserView(v)=!!v.userView` 作删除/改名保护判据。
- 视觉：用户视图名 `.vname.user`(强调色) + 首个用户视图前 `.view-sep` 分隔线；画册分组时去卡片顶色条(`if(!groupBy && galColor)`)。
- 右键菜单复用 `#ctxMenu` 容器(`showViewContextMenu` 仿 `showRecordContextMenu`)，重命名复用 `.overlay` 模态(仿 `showConfirm3`)。

## 用户技术背景
- VB.NET/COM、VSTO Excel 插件经验；C# 上手无障碍。
