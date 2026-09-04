#requires -Version 5.1

param(
    [ValidateSet('Auto', 'x86', 'x64')]
    [string]$Architecture = 'Auto',

    [string]$XllPath = ''
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'office_addin_setup.ps1')

$artifacts = Get-AddinArtifacts -ProjectName 'MultiTableAddin'
$detectedInfo = if ($Architecture -eq 'Auto') {
    Get-WpsBitness
}
else {
    [pscustomobject]@{
        Architecture = $Architecture
        Source       = 'Manual'
        Evidence     = '手动指定'
    }
}

$targetArchitecture = Resolve-TargetArchitecture -Architecture $Architecture -DetectedInfo $detectedInfo
$resolvedXllPath = if ([string]::IsNullOrWhiteSpace($XllPath)) {
    Resolve-AddinPathForArchitecture -Artifacts $artifacts -Architecture $targetArchitecture
}
else {
    [System.IO.Path]::GetFullPath($XllPath)
}

if (-not (Test-Path $resolvedXllPath)) {
    throw "XLL 文件不存在：$resolvedXllPath"
}

Register-WpsAddin -AddinPath $resolvedXllPath -Architecture $targetArchitecture -ProgIds @(
    'MultiTableAddin.Ribbon',
    'MultiTableAddin.CTPHost'
)

Write-InstallSummary `
    -HostKind 'WPS' `
    -Architecture $targetArchitecture `
    -Source $detectedInfo.Source `
    -Evidence $detectedInfo.Evidence `
    -XllPath $resolvedXllPath `
    -ActionSummary '已写入 WPS LoadMacros 和 AddinsWL 白名单'

$launchResult = Start-OfficeHost -HostKind 'WPS'
if ($launchResult.Started) {
    Write-Host ('已尝试启动 WPS：{0}' -f $launchResult.Path)
}
else {
    Write-Warning $launchResult.Message
}
