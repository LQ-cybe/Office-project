param(
    [string]$ProjectRoot = (Join-Path $PSScriptRoot '..'),
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [switch]$SkipBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-NormalizedPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return [System.IO.Path]::GetFullPath($Path)
}

function Assert-Condition {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-PathExists {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw ("缺少交付文件：{0} -> {1}" -f $Label, $Path)
    }
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    return ([System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json)
}

$projectRootPath = Resolve-NormalizedPath -Path $ProjectRoot
$buildScriptPath = Join-Path $projectRootPath 'build_addin.ps1'
$distRoot = Join-Path $projectRootPath 'dist'
$distFilesRoot = Join-Path $distRoot 'files'
$projectName = 'MultiTableAddin'

if (-not (Test-Path -LiteralPath $buildScriptPath)) {
    throw "未找到构建脚本：$buildScriptPath"
}

if (-not $SkipBuild) {
    & $buildScriptPath -Configuration $Configuration
    if ($LASTEXITCODE -ne 0) {
        throw 'build_addin.ps1 执行失败，已停止 dist 清单验收。'
    }
}

Assert-PathExists -Path $distRoot -Label 'dist 根目录'
Assert-PathExists -Path $distFilesRoot -Label 'dist/files 目录'

$requiredFiles = @(
    @{ Label = 'Excel 安装 bat'; Path = (Join-Path $distRoot 'install_excel_addin.bat') },
    @{ Label = 'WPS 安装 bat'; Path = (Join-Path $distRoot 'install_wps_addin.bat') },
    @{ Label = 'Excel 卸载 bat'; Path = (Join-Path $distRoot 'uninstall_excel_addin.bat') },
    @{ Label = 'WPS 卸载 bat'; Path = (Join-Path $distRoot 'uninstall_wps_addin.bat') },
    @{ Label = 'x86 XLL'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn.xll')) },
    @{ Label = 'x64 XLL'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn64.xll')) },
    @{ Label = 'x86 DNA'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn.dna')) },
    @{ Label = 'x64 DNA'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn64.dna')) },
    @{ Label = 'x86 deps'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn.deps.json')) },
    @{ Label = 'x64 deps'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn64.deps.json')) },
    @{ Label = 'x86 runtimeconfig'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn.runtimeconfig.json')) },
    @{ Label = 'x64 runtimeconfig'; Path = (Join-Path $distFilesRoot ($projectName + '-AddIn64.runtimeconfig.json')) },
    @{ Label = '主 DLL'; Path = (Join-Path $distFilesRoot ($projectName + '.dll')) },
    @{ Label = 'Excel 安装脚本'; Path = (Join-Path $distFilesRoot 'scripts\install_excel_addin.ps1') },
    @{ Label = 'WPS 安装脚本'; Path = (Join-Path $distFilesRoot 'scripts\install_wps_addin.ps1') },
    @{ Label = 'Excel 卸载脚本'; Path = (Join-Path $distFilesRoot 'scripts\uninstall_excel_addin.ps1') },
    @{ Label = 'WPS 卸载脚本'; Path = (Join-Path $distFilesRoot 'scripts\uninstall_wps_addin.ps1') }
)

foreach ($item in $requiredFiles) {
    Assert-PathExists -Path $item.Path -Label $item.Label
}

$runtimeConfig32Path = Join-Path $distFilesRoot ($projectName + '-AddIn.runtimeconfig.json')
$runtimeConfig64Path = Join-Path $distFilesRoot ($projectName + '-AddIn64.runtimeconfig.json')
$runtimeConfig32 = Read-JsonFile -Path $runtimeConfig32Path
$runtimeConfig64 = Read-JsonFile -Path $runtimeConfig64Path

Assert-Condition -Condition ($runtimeConfig32.runtimeOptions.tfm -eq 'net6.0') -Message 'x86 runtimeconfig 的 tfm 不是 net6.0。'
Assert-Condition -Condition ($runtimeConfig64.runtimeOptions.tfm -eq 'net6.0') -Message 'x64 runtimeconfig 的 tfm 不是 net6.0。'
Assert-Condition -Condition ($runtimeConfig32.runtimeOptions.rollForward -eq 'LatestPatch') -Message 'x86 runtimeconfig 的 rollForward 不是 LatestPatch。'
Assert-Condition -Condition ($runtimeConfig64.runtimeOptions.rollForward -eq 'LatestPatch') -Message 'x64 runtimeconfig 的 rollForward 不是 LatestPatch。'

Write-Host 'dist 交付清单校验通过'
Write-Host ('dist 目录：{0}' -f $distRoot)
Write-Host ('x86 XLL：{0}' -f (Join-Path $distFilesRoot ($projectName + '-AddIn.xll')))
Write-Host ('x64 XLL：{0}' -f (Join-Path $distFilesRoot ($projectName + '-AddIn64.xll')))
