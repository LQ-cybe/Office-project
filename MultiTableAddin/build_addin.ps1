#requires -Version 5.1

param(
    [string]$Configuration = 'Release'
)

$skillBuildScript = Join-Path $PSScriptRoot 'tools\build_addin.ps1'
if (-not (Test-Path $skillBuildScript)) {
    throw '未找到 tools\build_addin.ps1'
}

& $skillBuildScript -ProjectRoot $PSScriptRoot -Configuration $Configuration
