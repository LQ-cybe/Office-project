# MultiTableAddin WebView2 关闭生命周期验收清单

## 目标

确认 `WebView2 + CTP` 在下面几种路径下不会把 Excel 或 WPS 带崩：

- 关闭网页窗格
- 切换工作簿窗口后重复打开网页窗格
- 卸载插件
- 退出宿主进程

## 验收前准备

1. 正常安装插件并打开宿主。
2. 确认 `logs` 目录可写。
3. 在 Ribbon 中能看到 `网页窗格` 按钮。

## 用例 1：单次打开与关闭

1. 点击 `网页窗格`。
2. 在窗格中执行导航、刷新、A1 搜索。
3. 关闭窗格或触发窗格销毁。

预期：

- 宿主不崩溃。
- 再次打开窗格仍正常。
- 日志中能看到 `WebView2.Close.*` 相关记录。

## 用例 2：窗口切换与宿主差异

1. 在 Excel 中打开两个工作簿窗口；如果是 WPS，则打开两个工作簿并在同一 Application 内切换。
2. 在当前窗口打开 `网页窗格`。
3. 切换到另一个工作簿，再次观察或重新打开 `网页窗格`。
4. 来回切换并重复关闭、重新打开。

预期：

- Excel 下不出现全局单例窗格串窗。
- WPS 下应表现为同一 Application 内复用同一任务窗格，而不是伪造多套窗格实例。
- 日志中能解释当前宿主的窗格管理方式。

## 用例 3：插件卸载路径

1. 保持网页窗格处于已打开状态。
2. 触发插件卸载或重新加载。

预期：

- 宿主不闪退。
- 日志中至少能看到 `Lifecycle.AutoClose`。
- 如果不是宿主退出场景，允许出现 `WebView2.Close.Disposed`。

## 用例 4：宿主退出路径

1. 保持网页窗格处于已打开状态。
2. 直接关闭 Excel 或 WPS ET。

预期：

- 宿主退出过程不异常卡死或崩溃。
- 日志中应看到 `Lifecycle.ProcessExit`、`Lifecycle.DomainUnload` 或 `Lifecycle.AssemblyUnload` 之一。
- 这类路径下优先看到 `WebView2.Close.Deferred`，而不是强制重释放。

## 日志关注点

- `Lifecycle.AutoClose`
- `Lifecycle.ProcessExit`
- `Lifecycle.DomainUnload`
- `Lifecycle.AssemblyUnload`
- `WebView2.Close.Deferred`
- `WebView2.Close.Disposed`
- `WebView2.Close.Dispose.Error`

## 结论标准

- 4 个用例都不导致宿主崩溃
- 日志链路能解释关闭路径
- 重新打开窗格后仍可继续导航与交互
