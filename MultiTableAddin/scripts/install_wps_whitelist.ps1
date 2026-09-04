#requires -Version 5.1

param(
    [ValidateSet('Auto', 'x86', 'x64')]
    [string]$Architecture = 'Auto',

    [string]$XllPath = ''
)

$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'install_wps_addin.ps1') -Architecture $Architecture -XllPath $XllPath
