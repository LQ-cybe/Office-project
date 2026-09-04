#requires -Version 5.1

param(
    [string]$XllPath = ''
)

$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'uninstall_wps_addin.ps1') -XllPath $XllPath
