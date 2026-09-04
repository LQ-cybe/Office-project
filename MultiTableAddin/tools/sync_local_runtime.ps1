#requires -Version 5.1

param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [switch]$DryRun,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Copy-DirectoryContent {
    param(
        [string]$SourceRoot,
        [string]$TargetRoot,
        [switch]$ForceCopy
    )

    $files = Get-ChildItem -Path $SourceRoot -Recurse -File
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $targetPath = Join-Path $TargetRoot $relative
        $parent = Split-Path -Parent $targetPath

        if ($DryRun) {
            Write-Host ("[复制] {0}" -f $relative)
            continue
        }

        if ($parent -and -not (Test-Path $parent)) {
            [void][System.IO.Directory]::CreateDirectory($parent)
        }

        if ((Test-Path $targetPath) -and -not $ForceCopy) {
            continue
        }

        [System.IO.File]::Copy($file.FullName, $targetPath, $true)
    }
}

$skillRoot = Split-Path -Parent $PSScriptRoot
$runtimeSource = Join-Path $skillRoot 'runtime\exceldna'
$projectRuntime = Join-Path ([System.IO.Path]::GetFullPath($ProjectRoot)) 'runtime\exceldna'
$resolvedRuntimeSource = [System.IO.Path]::GetFullPath($runtimeSource)
$resolvedProjectRuntime = [System.IO.Path]::GetFullPath($projectRuntime)

if (-not (Test-Path $runtimeSource)) {
    throw "未找到本地运行时目录: $runtimeSource"
}

if ([string]::Equals($resolvedRuntimeSource, $resolvedProjectRuntime, [System.StringComparison]::OrdinalIgnoreCase)) {
    if ($DryRun) {
        Write-Host "预演同步 runtime：源目录与目标目录相同，跳过复制。"
    }
    else {
        Write-Host "本地运行时已就位，无需重复同步：$resolvedProjectRuntime"
    }
    return
}

if ($DryRun) {
    Write-Host "预演同步 runtime：$runtimeSource -> $projectRuntime"
}
else {
    if ($Force -and (Test-Path $projectRuntime)) {
        [System.IO.Directory]::Delete($projectRuntime, $true)
    }
    [void][System.IO.Directory]::CreateDirectory($projectRuntime)
}

Copy-DirectoryContent -SourceRoot $runtimeSource -TargetRoot $projectRuntime -ForceCopy:$Force

if (-not $DryRun) {
    Write-Host "本地运行时已同步：$projectRuntime"
}

<#
dist_id: 3a8864a2-56f90dc3-54ef559a-0fbf3b47-811e811e-80d7801a-a96ff00a
ctx_id: b610b0d1-508c1c35-0b8c5542-37f50d43-53951c35-0a87555e-0130f5a9-d575dc34-34bc555d-20f53953-96f63951-508c393e-0a9c546a-33f92941-509e3837-2b935745-1ef63866-5e971a36-22b85f6d-3af71650-50bd1235-0dab546c-23f50d73-53ac3f36-2c945642-33f8377b-52ac1037-24bd5351-34f51f68-52aa3e37-37a65655-39f40c71-50821d32-36915841-13f53878-50901735-0ab05643-1bf8115d-52a80a3e-0a9c5659-27f62668-53a03635-098d5744-2ff80f6c-51b90637-0585556f-3df80472-52ab0b36-2c94564c-35f53878-559032
#>
