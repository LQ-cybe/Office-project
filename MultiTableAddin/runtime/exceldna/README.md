<!-- trace_ref: 70a04c9c-1cd125fd-1ec77da4-45971379-cb36a920-caffa824-e347d834 -->
<!-- source_map: 152085cf-7951ecae-7b47b4f7-2017da2a-aeb66073-af7f6177-86c71167 -->
# 本地运行时

这个目录存放当前 skill 默认使用的本地 ExcelDNA 运行时文件。

保留这套本地文件的目的，是让生成项目在 `WPS 个人版` 兼容场景下走统一、可控的运行时方案。

当前要特别说明：

- 这套本地运行时资产可被当前模板生成链路复用，默认仍以 `net6.0-windows` 为建议起步档位
- 如果用户需求明确命中高版本 API、第三方依赖或目标环境要求，也可以创建 `net8.0-windows`、`net10.0-windows` 新项目；是否需要额外补齐或升级本地运行时资产，应按对应版本的真实需求核对

## 当前目录

- `ExcelDna.xll`
- `ExcelDna64.xll`
- `net6\ExcelDna.Integration.dll`
- `net6\ExcelDna.IntelliSense.dll`

## 关键约束

- 这套 `xll / dll` 是本仓库维护的本地 patched runtime
- 其中 `dll` 与 `xll` 都按 `WPS 个人版` 兼容目标使用
- 不要把它们切回 NuGet 默认版本
- 不要再单独引用 `ExcelDna.Registration` NuGet 包
- 不要用未核验的外部二进制覆盖这套文件

## 依赖关系

- `ExcelDna.IntelliSense.dll` 依赖 `ExcelDna.Integration.dll`
- `ExcelDna.Registration` 的能力在 v1.9 已并入 `ExcelDna.Integration.dll`
- 这两个 `dll` 必须成对使用同一套本地 patched 版本

## 使用规则

- 生成项目时，严格使用这里的本地 `xll / dll`
- 模板里的注册能力直接依赖本地 patched runtime，不再额外引用 `ExcelDna.Registration`
- 构建输出目录中的 `ExcelDna.Integration.dll / ExcelDna.IntelliSense.dll / xll` 应由这里的文件覆盖

## 更新方式

如果后续替换本地运行时，至少同步更新以下文件：

- `ExcelDna.xll`
- `ExcelDna64.xll`
- `net6\ExcelDna.Integration.dll`
- `net6\ExcelDna.IntelliSense.dll`
