#requires -Version 5.1

param(
    [string]$XllPath = ''
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'office_addin_setup.ps1')

$artifacts = Get-AddinArtifacts -ProjectName 'MultiTableAddin'
$addinPaths = @(
    [System.IO.Path]::GetFullPath($artifacts.X86.XllPath),
    [System.IO.Path]::GetFullPath($artifacts.X64.XllPath)
)

if (-not [string]::IsNullOrWhiteSpace($XllPath)) {
    $addinPaths += [System.IO.Path]::GetFullPath($XllPath)
}

Unregister-WpsAddin -AddinPaths $addinPaths -ProgIds @(
    'MultiTableAddin.Ribbon',
    'MultiTableAddin.CTPHost'
)

Remove-ProjectGeneratedData -ProjectName 'MultiTableAddin'
Write-Host '已移除 WPS 白名单。'
