# MultiTableAddin 蓝图说明

这个蓝图专门演示 `WebView2 + CTP` 在 Excel/WPS 宿主里的完整接入架构。  
它不仅覆盖关闭阶段，还覆盖初始化、事件绑定、导航、宿主退出和释放收口。凡是项目里要接 `WebView2`，都应优先沿用这套架构，而不是只摘一段控件初始化代码。

> 这个蓝图已经主动移除了异步专项示例。  
> 如果需求重点只是 `公式异步`、`命令异步` 或 `WPS` 下的异步分支处理，先去 `templates/project-blueprints/wps-async-formula-command/`。  
> 只有当需求明确包含 `WebView2` 时，才进入这个蓝图。
> 如果要改 `Ribbon.xml`、补 `getImage / onAction / onLoad` 等回调，先看：`references/22-exceldna-ribbon-callback-signatures.md`

## 适用场景

- 你需要 `Ribbon + CTP + WPF/HandyControl` 的完整工作台
- 你需要在 Excel/WPS 插件里接入 `WebView2`，并希望从一开始就按可排障的生命周期架构落地
- 需要区分“插件卸载”和“Excel 进程退出”两种清理路径
- 需要让 CTP 销毁、ElementHost 脱钩、WebView2 事件解绑更可控

## 源码结构

- `src/app_name/AddIn.cs.tmpl`
  - 只放插件生命周期、日志、宿主识别和运行时初始化
- `src/app_name/TaskPane/`
  - 放 CTP 宿主控件、WPF 视图和 `WebView2` 生命周期收口
- `src/app_name/RibbonController.cs.tmpl`、`WorkbookCommands.cs.tmpl`
  - 放 Ribbon 分发与工作簿自动化命令
- `src/app_name/Functions.cs.tmpl`、`Views/`、`Resources/`
  - 放最小 UDF、辅助视图和 Ribbon 资源

## 内置能力

- 保留完整工作台所需的基础 UDF、Ribbon、CTP、多窗口管理能力
- 新增 `ExcelWorkbookDataOps`：
  - 用 `Range.Value2` 批量读写二维数组
  - 用 `ExcelAppStateScope` 统一保存和恢复 `ScreenUpdating / EnableEvents / DisplayAlerts / Calculation / StatusBar`
  - 用 `ListObjectPreserveSpec` 安全重写结构化表，并保留公式列与列格式
  - 用列 schema 统一处理文本列、日期列、日期时间列、时间列、金额列
- 新增 `WorkbookCommands` 的工作簿自动化样板：
  - `导出表格`：新工作表 + 整块写入 + 转 `ListObject`
  - `重写表格`：保留公式列和金额格式的原位重写
- 增加 `AddInLifecycle`：记录 `AutoClose / ProcessExit / DomainUnload / AssemblyUnload`
- 增加 `TaskPaneCloseContext`：把关闭来源和清理策略传给窗格
- `WebTaskPaneWpfView` 的所有初始化、导航、搜索、事件绑定与关闭都统一纳入生命周期约束
- `WebTaskPaneHostControl.Dispose()` 先通知 WPF 视图准备关闭，再脱离宿主
- `WebTaskPaneWpfView` 在宿主关闭阶段只解绑事件、脱离 `ElementHost`，避免激进 `Dispose`
- 窗格绑定模型要区分宿主：Excel 2013+ 按活动窗口管理；WPS 个人版通常是单 Application 共用同一任务窗格

## 推荐点击路径

1. 打开 Excel 或 WPS ET，点击 `网页窗格`
2. 在窗格中执行几次导航、刷新、A1 搜索，确认 WebView2 正常工作
3. 关闭窗格、切换工作簿窗口、再次打开窗格，观察是否仍正常
4. 退出宿主或卸载插件后，检查日志中的 `Lifecycle.*` 与 `WebView2.Close.*`

详细验收步骤见 `docs/webview2-close-checklist.md`

## 学习重点

- `AutoClose` 不是 Excel 真正退出的唯一信号
- 在 ExcelDNA 中实现工作簿自动化时，优先延续“读入数组 -> 内存计算 -> 整块写回”的主线，不要把逐格 COM 访问直接写回 C# 里
- 结构化明细输出优先 `ListObject`，原位重写时先保留公式列和格式列，再批量写回
- 长编号、日期、时间、日期时间、金额列要先做列 schema 设计，再决定写入和格式策略
- 所有 `WebView2` 使用都要放在这套生命周期架构里管理，不能只复制初始化代码，把释放逻辑留到以后补
- 宿主正在退出时，不要强制把 `WebView2` 做重释放
- 先解绑事件、清空宿主引用，再决定是否调用 `Dispose`
- WebView2 问题属于宿主生命周期与资源管理问题，不该和普通异步样板混在一起

## AI 使用建议

- 当前模板仓库里，普通工作台和 `WebView2` 工作台都从这个蓝图起步
- 如果用户最终不需要 `WebView2`，在生成后的项目里删掉网页窗格、相关引用和收口日志即可
- 如果用户最终不需要网页窗格，不要把这套 `WebView2` 生命周期代码机械带入正式项目

## 冒烟建议

- 自动冒烟现在会额外验证：新工作表导出、`ListObject` 安全重写、长编号文本格式、日期时间格式和公式列恢复
- 人工重点验证 `网页窗格` 的打开、关闭、切换窗口、退出宿主
- 查看 `logs` 中是否按顺序出现 `Lifecycle.*`、`WebView2.Close.Deferred`、`WebView2.Close.Disposed`

## 最终用户示例文件

- 冒烟通过后，AI 应补一份贴合最终用户真实工作台功能的测试文件，而不是只停留在模板自带的网页窗格演示与工作簿自动化样板
- 如果项目重点是网页检索、任务窗格交互、导入导出、结构化表写回或宿主联动，就应让示例文件直接体现这些入口、参数和预期结果
- 示例文件应帮助用户复测 `UDF`、Ribbon 命令、任务窗格配合关系，以及关键输入输出工作表，不应只是开发者用来观察日志的空壳 demo
- 执行 `build_addin.ps1` 前，应确认示例文件已经围绕当前正式功能重建；交付时再把它一并放进 `dist/examples/`
