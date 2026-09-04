# 多维表分析软件 — 本项目 Skill 使用分析与优化建议

> 编写时间：2026-08-14  
> 范围：复盘本项目（Excel-DNA + .NET 8 + WebView2 单文件 SPA 插件）开发过程中实际调用的 Skill，评估其价值与不足，并给出可落地的优化建议。

---

## 一、本项目实际使用的 Skill

| Skill | 用途 | 使用频次 | 价值 |
|-------|------|----------|------|
| `exceldna-build-verify` | 构建→交付→验证流水线；`verify-build.py` 字节子串校验；无 SVG 渲染器时的图标栅格化知识 | 极高（每次改动必走） | ★★★★★ |
| `exceldna-addin-builder` | Excel-DNA / XLL / Ribbon / CTP / WPF / HandyControl 架构与最佳实践 | 高（架构与功能开发） | ★★★★☆ |
| `icon-design` | 视图/功能区图标检索（Tabler 等）、SVG 回传、导出 | 中（多次重选图标 #158/#193/#216/#234/#273） | ★★★★☆ |
| `xlsx` | 测试数据表格的读取/生成辅助 | 低 | ★★★☆☆ |

> 其它 Excel/Office 类 Skill（`Excel自动化`、`tencent-docs`、`tencent-local-office-edit`、`pdf`）在本项目中未实际触发——本插件的数据源是 Excel 超级表，由 C# 侧直接经 Excel-DNA 读写，未走这些 Skill 的自动化路径。

---

## 二、各 Skill 的使用方式与价值

### 1. exceldna-build-verify（核心依赖）
- **提供了**：`dotnet build -c Release` → 拷贝 DLL 到 `dist/files` → `python verify-build.py` 的完整流水线；UTF-16LE / UTF-8 双编码校验规则；无本机 SVG 渲染器时栅格化的思路。
- **价值**：把「改了没生效」「校验 FAIL」这类高频坑固化成了可重复流程，是本项目能稳定迭代 374 项任务的基础设施。
- **实际卡点**：流水线本身正确，但 skill 未强制强调「build 产物路径 `src/.../bin/Release/net8.0-windows/` 与安装取用路径 `dist/files/` 不同」，新手极易只 build 不重拷。

### 2. exceldna-addin-builder（架构底座）
- **提供了**：Excel-DNA 1.9 本地 patched runtime、`net8.0-windows`、WPF+HandyControl、WebView2 内嵌 `index.html` 的桥接模式（`window.api.call` ↔ `HtmlMainForm.Dispatch`）。
- **价值**：开箱即用的架构骨架，省去反复试错。
- **实际卡点**：本项目同时 `UseWindowsForms`+`UseWPF`，出现 `Rectangle`/`Color` 类型歧义（CS1729/CS1503），skill 未单列此坑。

### 3. icon-design（图标选型）
- **提供了**：按关键词检索图标库（Tabler 等）、返回 SVG、导出成品图。
- **价值**：多次快速替换视图/功能区图标，最终统一为 Tabler 线性风格。
- **实际卡点**：检索结果是「单个图标」，而本项目需要「批量汇总已用图标并出对照表」，skill 无此聚合能力；且返回 SVG 含 `currentColor` 特性，文档未提示「可直接改 stroke 颜色复用」。

---

## 三、Skill 使用中的痛点与不足

1. **build-verify 未固化「两步交付」检查清单**：只说「构建→拷贝→校验」，但没强调「只 build 不重拷 dist = 旧资源仍被加载」，这是本项目最高频的隐性失误。
2. **build-verify 未收录 WPF/WinForms 类型歧义**：绘图代码（导出裁剪）必须用 `System.Drawing.*` 全限定名，否则编译不过。
3. **build-verify 的 verify-build 编码规则可再显眼**：「C# 字符串 UTF-16LE / JS UTF-8 / 注释不进 DLL」是校验失败头号原因，建议放在流水线最前作为「铁律」。
4. **icon-design 缺「批量汇总/导出对照表」能力**：本项目最终要「把用过的图标汇总 + 出对照表」，只能自写脚本实现，skill 未覆盖。
5. **icon-design 的 SVG 复用说明不足**：`currentColor` 可随主题变色这一关键特性未在返回说明里点出，导致初期误以为要改色就得重新生成。
6. **缺少「无第三方库的最小实现」范式**：手写 PDF（FlateDecode + Predictor 15 + ZLibStream）曾是本项目硬骨头，现有 Skill 未提供这类「零依赖兜底」模板（#376 已移除 PDF 导出功能，但经验可沉淀）。

---

## 四、Skill 优化建议（可落地）

### 对 `exceldna-build-verify` 的建议
- **A. 新增「交付两步法」强制清单**：在流水线最前写明——① `dotnet build -c Release`（产物在 `src/.../bin/Release/net8.0-windows/`）；② 拷贝 `MultiTableAddin.dll` 等到 `dist/files/`（覆盖前先 `rm -f` 防句柄锁）；③ `python verify-build.py`。明确「只做①不算交付」。
- **B. 把编码铁律置顶**：`verify-build.py` 必须以对应编码搜索——C# 字面量 `utf-16le`、内嵌 HTML/JS `utf-8`；**注释/源码片段不进 DLL，禁止用作校验串**。每次改 `index.html` 必同步校验串（正向+负向）。
- **C. 增加「WPF/WinForms 混用」提示卡**：绘图相关类型一律 `System.Drawing.Rectangle` / `System.Drawing.Color.White` / `System.Drawing.Imaging.*` 全限定。
- **D. 增补「WebView2 截图裁剪」片段**：JS `getBoundingClientRect()` 取视口坐标 → C# 用 `_webView.ClientSize` 比例映射到 Bitmap 像素 → `Bitmap.Clone(Rectangle,...)`；底部以内容首子元素 `bottom` 为基准。
- **E. 增补「零依赖 PDF」模板（经验沉淀）**：DeviceRGB + `/Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors 3 /BitsPerComponent 8 /Columns w >>`，zlib 用 `System.IO.Compression.ZLibStream`；先铺白底转 24bit RGB 防透明变黑；1px=1pt。> 注：#376 已移除 PDF 导出（位图为栅格、非矢量），该模板仅作经验沉淀，若未来需矢量 PDF 再实现时可复用。

### 对 `exceldna-addin-builder` 的建议
- **F. 标注 dist 与 src 路径差异**：明确 `dist/files/` 才是安装取用目录，`src/.../bin/Release/` 只是编译中间物。
- **G. 说明 Ribbon 实际启用项**：本项目的 `Ribbon.xml` 仅 `btnOpenMultiTable`（tag=`RibbonOpen`）真正启用，其余 `RibbonActivity/Order/About` 仅内嵌待用——避免误以为要多按钮联动。

### 对 `icon-design` 的建议
- **H. 新增「批量汇总已用图标」模式**：输入一组图标 key/文件，输出独立 `.svg` + 元数据 `.svg.json` + 对照表 MD（本项目已自写 `extract_icons.py` 实现，可回流为 skill 能力）。
- **I. 强化 SVG 复用说明**：返回时标注「`stroke=currentColor` 可跟随父元素 `color` 变色；去 `width/height` 后可无损缩放」；并提示批量去属性时正则须锚定 `\swidth=` 以免误伤 `stroke-width`。

### 建议新增的 Skill
- **`exceldna-icon-kit`（或在 build-verify 内增节）**：把「图标提取 + 多尺寸栅格化 + 对照表生成」打包，覆盖本项目的 #374 需求，避免每次手搓脚本。

---

## 五、本项目沉淀的可复用资产

| 资产 | 位置 | 可回流为 Skill 能力 |
|------|------|-------------------|
| 图标提取+对照表脚本 | `temp/extract_icons.py` | → icon-design「批量汇总」模式 |
| verify-build 双编码校验范式 | `verify-build.py` | → build-verify 编码铁律 |
| 零依赖 PDF 生成片段（#376 已移除） | 原 `HtmlMainForm.cs` `SaveBitmapAsPdf` | → build-verify「零依赖 PDF」模板（经验沉淀） |
| WebView2 裁剪片段 | `HtmlMainForm.cs` + `index.html getExportCrop()` | → build-verify「截图裁剪」片段 |
| 构建两步法 | 项目根 `tools/build_addin.ps1` | → build-verify 交付清单 |

> 以上资产均可回流进对应 Skill，使下一次同类项目直接复用，无需重新踩坑。
