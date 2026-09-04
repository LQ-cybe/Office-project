#requires -Version 5.1
<#
开发调试辅助脚本：把 Excel 加载的插件在「Debug 构建输出」与「正式 dist 版本」之间切换。

用法：
    .\dev_switch_debug_addin.ps1            # 切到 Debug 构建（配合 VS 按 F5 调试）
    .\dev_switch_debug_addin.ps1 -Target prod   # 恢复正式安装版本

原理：ExcelDNA 插件通过注册表 OPEN 项被 Excel 加载。
调试时把注册路径指向 src\MultiTableAddin\bin\Debug 下的 XLL，
Visual Studio F5 启动 Excel 时即加载最新 Debug 代码，可直接命中断点。
#>

param(
    [ValidateSet('debug', 'prod')]
    [string]$Target = 'debug'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent
. (Join-Path $root 'scripts\office_addin_setup.ps1')

$distXlls = @(
    (Join-Path $root 'dist\files\MultiTableAddin-AddIn64.xll'),
    (Join-Path $root 'dist\files\MultiTableAddin-AddIn.xll'),
    (Join-Path $root 'dist\files_wpf\MultiTableAddin-AddIn64.xll'),
    (Join-Path $root 'dist\files_wpf\MultiTableAddin-AddIn.xll')
)
$debugDir = Join-Path $root 'src\MultiTableAddin\bin\Debug\net8.0-windows'
$debugXlls = @(
    (Join-Path $debugDir 'MultiTableAddin-AddIn64.xll'),
    (Join-Path $debugDir 'MultiTableAddin-AddIn.xll')
)

$existingDist = $distXlls | Where-Object { Test-Path $_ }
$existingDebug = $debugXlls | Where-Object { Test-Path $_ }

# 先清理所有可能的注册，避免 Excel 同时加载两份
if ($existingDist) { Unregister-ExcelAddinRegistry -AddinPaths $existingDist }
if ($existingDebug) { Unregister-ExcelAddinRegistry -AddinPaths $existingDebug }

if ($Target -eq 'prod') {
    if (-not $existingDist) { throw '未找到 dist 安装产物，请先运行 install_excel_addin.bat 或 build_addin.ps1' }
    Register-ExcelAddinRegistry -AddinPath ($existingDist | Select-Object -First 1)
    Write-Host '已切换到：正式 dist 版本'
} else {
    if (-not $existingDebug) {
        Write-Host 'Debug 构建不存在，正在执行 dotnet build ...' -ForegroundColor Yellow
        dotnet build (Join-Path $root 'src\MultiTableAddin\MultiTableAddin.csproj') -c Debug
        if ($LASTEXITCODE -ne 0) { throw 'dotnet build 失败' }
        $existingDebug = $debugXlls | Where-Object { Test-Path $_ }
    }
    # Excel 64 位优先加载 64 位 XLL
    $pick = if (Test-Path $debugXlls[0]) { $debugXlls[0] } else { $debugXlls[1] }
    Register-ExcelAddinRegistry -AddinPath $pick
    Write-Host "已切换到：Debug 构建（$pick）"
}

Write-Host ''
Write-Host '现在可以在 Visual Studio 中打开 src\MultiTableAddin\MultiTableAddin.csproj，'
Write-Host '在调试下拉框选择「启动 Excel 调试插件」，按 F5 即可带调试器启动 Excel。'
Write-Host '注意：运行前请确保已完全关闭所有 Excel 窗口。'

