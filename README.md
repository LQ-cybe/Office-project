# 多维表分析软件（MultiTableAddin）

> 一款运行在 **Microsoft Excel / WPS 表格** 内的多维数据分析插件：把一张普通的 Excel 超级表（`Ctrl+T`），瞬间变成支持 **表格、表单、画册、日历、甘特图、仪表盘、图表** 七大视图的多维数据工作台。

- **当前版本**：v2.9.42
- **开源协议**：[MIT License](MultiTableAddin/LICENSE)
- **技术底座**：ExcelDNA 1.9 + .NET 8 + WebView2 + WPF/HandyControl
- **双宿主支持**：Microsoft Excel（Office 2016 及以上）与 WPS 表格共用同一套 DLL/XLL，自动识别宿主

---

## 📁 仓库结构

本仓库为项目完整备份，包含插件源码、全部开发文档、更新记录、测试数据与参考资料。

| 目录/文件 | 说明 |
|-----------|------|
| [`MultiTableAddin/`](MultiTableAddin/README.md) | ★ 插件源码工程（含源码构建与安装指引，详见其 README） |
| [`文档中心/`](文档中心/README.md) | 全部项目文档：使用手册、功能设计、规划方案、开发记录、早期归档 |
| `更新记录_甘特对齐图标配色与按钮更名.md` | 更新记录（2026-09-01 补丁：功能区更名 + 甘特对齐修复 + 图标颜色统一） |
| `窗口模式与最小化浮窗分析.md` | 窗口模式与「最小化浮窗」特性分析 |
| `测试数据/` | 测试用 xlsx / multiview.json、测试图片、数据生成脚本 |
| `脚本与工具/` | 构建验证脚本、JS 语法检查、图标重着色工具 |
| `参考资料/` | 竞品对比、甘特/日历交互参考资料、GanttWPF 示例、图标素材 |
| `.workbuddy/` | 开发过程记忆与历史快照备份 |

### 文档中心子目录

| 目录 | 内容 |
|------|------|
| `01-使用手册` | 面向使用者的功能说明 |
| `02-功能设计` | 各功能模块的设计与实现分析 |
| `03-规划方案` | 技术选型与未来规划 |
| `04-开发记录` | 项目开发记录、版本更新记录、经验总结 |
| `99-归档_早期草稿` | 已被正式版文档取代的早期草稿（仅留档） |

---

## 🚀 快速开始（使用者）

1. 确认已安装 **.NET 8 Desktop Runtime**（Windows 10/11）
2. 完全关闭 Excel/WPS
3. 构建后进入 `MultiTableAddin\dist\` 目录：
   - Excel 用户双击 `install_excel_addin.bat`
   - WPS 用户双击 `install_wps_addin.bat`
4. 打开 Excel/WPS，将数据区域按 `Ctrl+T` 转换为超级表
5. 点击 Ribbon 上的「多维视图」按钮即可打开主窗口

> 详细安装/卸载说明见 [`MultiTableAddin/安装说明.txt`](MultiTableAddin/安装说明.txt)。

---

## 🛠️ 源码构建（开发者）

### 环境要求

- Windows 10/11
- Visual Studio 2022 或 .NET 8 SDK
- 仓库内已附带本地补丁版 ExcelDNA 运行时（`runtime\exceldna\`），构建强依赖，请勿删除

### 构建流程

```powershell
cd MultiTableAddin
.\build_addin.ps1 -Configuration Release
```

> ⚠️ 仅执行 `dotnet build` 不会更新安装目录，必须运行 `build_addin.ps1`；构建后需完全关闭 Excel/WPS 再重新运行安装脚本。

---

## 📌 版本信息

- **v2.9.42**（2026-08-20）：看板能力并入画册分组拖拽；图标统一强调蓝；表格列宽自适应、表单复制新增、分组折叠；视图右键菜单与配置保存位置开关等多项打磨
- **2026-09-01 补丁**：功能区按钮「多维表 → 多维视图」更名、甘特月/年视图右侧对齐 BUG 修复、功能图标统一强调蓝（详见根目录更新记录）

详细开发历程见 `文档中心/04-开发记录/`。

---

## 📄 许可证

本项目基于 [MIT License](MultiTableAddin/LICENSE) 开源，欢迎使用、修改与分发。
