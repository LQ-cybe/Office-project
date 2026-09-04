#requires -Version 5.1

param(
    [string]$ProjectRoot = '',

    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'

function Assert-DotnetAvailable {
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -eq $dotnet) {
        throw '未检测到 dotnet CLI，请先安装 .NET SDK。'
    }
}

function Invoke-DotnetCommand {
    param(
        [string[]]$Arguments,
        [string]$WorkingDirectory
    )

    Push-Location $WorkingDirectory
    try {
        Write-Host ('执行命令: dotnet ' + ($Arguments -join ' '))
        & dotnet @Arguments
        if ($LASTEXITCODE -ne 0) {
            throw "命令执行失败：dotnet $($Arguments -join ' ')"
        }
    }
    finally {
        Pop-Location
    }
}

function Copy-DirectoryArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$TargetRoot
    )

    if (-not (Test-Path $SourceRoot)) {
        return
    }

    Get-ChildItem -Path $SourceRoot -Recurse -File | ForEach-Object {
        $relativePath = $_.FullName.Substring($SourceRoot.Length).TrimStart('\')
        $targetPath = Join-Path $TargetRoot $relativePath
        $targetParent = Split-Path -Parent $targetPath
        if (-not (Test-Path $targetParent)) {
            [void][System.IO.Directory]::CreateDirectory($targetParent)
        }

        try {
            [System.IO.File]::Copy($_.FullName, $targetPath, $true)
        }
        catch [System.IO.IOException] {
            Write-Warning ("目标文件被占用，已跳过复制：{0}" -f $targetPath)
            Write-Warning '请先关闭正在加载该 XLL 的 Excel/WPS 进程后重新执行 build_addin.ps1。'
        }
    }
}

function Copy-FileArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$TargetRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )

    if (-not (Test-Path $SourceRoot)) {
        return
    }

    foreach ($pattern in $Patterns) {
        Get-ChildItem -Path $SourceRoot -Filter $pattern -File -ErrorAction SilentlyContinue | ForEach-Object {
            [System.IO.File]::Copy($_.FullName, (Join-Path $TargetRoot $_.Name), $true)
        }
    }
}

function Invoke-ExampleWorkbookGenerator {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$ProjectScriptsDir
    )

    $generatorScriptPath = Join-Path $ProjectScriptsDir 'generate_bi_demo_workbook.ps1'
    if (-not (Test-Path $generatorScriptPath)) {
        return
    }

    Write-Host ("检测到示例工作簿脚本，开始生成测试文件：{0}" -f $generatorScriptPath)
    & $generatorScriptPath -ProjectRoot $ProjectRoot
    if ($LASTEXITCODE -ne 0) {
        throw 'generate_bi_demo_workbook.ps1 执行失败，未能生成示例工作簿。'
    }
}

function Write-TextFileUtf8Bom {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path $parent)) {
        [void][System.IO.Directory]::CreateDirectory($parent)
    }

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Content.TrimEnd() + "`r`n", $utf8Bom)
}

function Write-TextFileAscii {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path $parent)) {
        [void][System.IO.Directory]::CreateDirectory($parent)
    }

    $ascii = New-Object System.Text.ASCIIEncoding
    [System.IO.File]::WriteAllText($Path, $Content.TrimEnd() + "`r`n", $ascii)
}

function Get-OfficeLaunchScriptContent {
    return @'
#requires -Version 5.1

param(
    [ValidateSet('Excel', 'WPS')]
    [string]$HostKind
)

$ErrorActionPreference = 'Stop'

function Get-RegistryStringValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$SubKey,
        [AllowEmptyString()]
        [string]$Name = ''
    )

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        try {
            if ($null -eq $key) {
                return $null
            }

            return [string]$key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        }
        finally {
            if ($null -ne $key) {
                $key.Dispose()
            }
        }
    }
    finally {
        if ($null -ne $baseKey) {
            $baseKey.Dispose()
        }
    }
}

function Get-ExcelExecutablePath {
    $candidates = @(
        (Get-RegistryStringValue -Hive LocalMachine -View Registry64 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'),
        (Get-RegistryStringValue -Hive LocalMachine -View Registry32 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'),
        (Get-RegistryStringValue -Hive CurrentUser -View Registry64 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'),
        (Get-RegistryStringValue -Hive CurrentUser -View Registry32 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe')
    )

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return ''
}

function Get-WpsExecutablePath {
    $candidates = @(
        (Get-RegistryStringValue -Hive CurrentUser -View Registry64 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\et.exe'),
        (Get-RegistryStringValue -Hive CurrentUser -View Registry32 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\et.exe'),
        (Get-RegistryStringValue -Hive LocalMachine -View Registry64 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\et.exe'),
        (Get-RegistryStringValue -Hive LocalMachine -View Registry32 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\et.exe')
    )

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $overlay = Get-RegistryStringValue -Hive CurrentUser -View Registry32 -SubKey 'Software\Classes\SystemFileAssociations\.et' -Name 'TypeOverlay'
    if (-not [string]::IsNullOrWhiteSpace($overlay)) {
        $office6Marker = '\office6'
        $index = $overlay.ToLowerInvariant().IndexOf($office6Marker)
        if ($index -ge 0) {
            $wpsPath = $overlay.Substring(0, $index + $office6Marker.Length) + '\et.exe'
            if (Test-Path $wpsPath) {
                return $wpsPath
            }
        }
    }

    return ''
}

$processPath = if ($HostKind -eq 'Excel') { Get-ExcelExecutablePath } else { Get-WpsExecutablePath }
if ([string]::IsNullOrWhiteSpace($processPath)) {
    Write-Warning ("未找到 {0} 可执行文件，请手动启动。" -f $HostKind)
    exit 1
}

Start-Process -FilePath $processPath | Out-Null
Write-Host ("已尝试启动 {0}：{1}" -f $HostKind, $processPath)
'@
}

function Get-InstallStatusScriptContent {
    return @'
#requires -Version 5.1

param(
    [ValidateSet('Success', 'Failure')]
    [string]$Result,
    [ValidateSet('Excel', 'WPS')]
    [string]$HostKind,
    [string]$LogPath = ''
)

$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '========================================'
if ($Result -eq 'Success') {
    Write-Host ($HostKind + ' 加载项安装成功')
} else {
    Write-Host ($HostKind + ' 加载项安装失败')
}
Write-Host '========================================'

if ($Result -eq 'Success') {
    Write-Host ('宿主：{0}' -f $HostKind)
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        Write-Host ('安装日志：{0}' -f $LogPath)
    }
    Write-Host ('已尝试自动打开 {0}，如未弹出请手动启动。' -f $HostKind)
    Write-Host '如需排查问题，可查看上面的安装日志文件。'
} else {
    Write-Warning '安装脚本执行失败，请先查看安装日志。'
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        Write-Host ('安装日志：{0}' -f $LogPath)
    }
}
'@
}

function Get-BatLauncherContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$ScriptRelativePath,

        [Parameter(Mandatory = $true)]
        [string]$LogFileName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Excel', 'WPS')]
        [string]$StatusHostKind,

        [string]$LaunchHostKind = ''
    )

    $launchBlock = ''
    if (-not [string]::IsNullOrWhiteSpace($LaunchHostKind)) {
        $launchBlock = @'
  "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0files\scripts\launch_office_host.ps1" -HostKind __HOST_KIND__ >> "%LOG_PATH%" 2>&1
'@.Replace('__HOST_KIND__', $LaunchHostKind)
    }

    return @'
@echo off
setlocal
title __DISPLAY_NAME__
set "SCRIPT_PATH=%~dp0__SCRIPT_RELATIVE_PATH__"
set "LOG_DIR=%~dp0files\logs"
set "LOG_PATH=%LOG_DIR%\__LOG_FILE_NAME__"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%" >nul 2>nul
echo ========================================
echo __DISPLAY_NAME__
echo ========================================
echo.
if not exist "%SCRIPT_PATH%" (
  echo ERROR: script not found
  echo   %SCRIPT_PATH%
  echo.
  pause
  exit /b 1
)
echo Running installer script:
echo   %SCRIPT_PATH%
echo Log file:
echo   %LOG_PATH%
echo.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" %* > "%LOG_PATH%" 2>&1
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (
__LAUNCH_BLOCK__
  "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0files\scripts\show_install_status.ps1" -Result Success -HostKind __HOST_KIND_FOR_STATUS__ -LogPath "%LOG_PATH%"
) else (
  "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0files\scripts\show_install_status.ps1" -Result Failure -HostKind __HOST_KIND_FOR_STATUS__ -LogPath "%LOG_PATH%"
  echo Exit code: %EXITCODE%
)
echo.
pause
exit /b %EXITCODE%
'@.Replace('__DISPLAY_NAME__', $DisplayName).
        Replace('__SCRIPT_RELATIVE_PATH__', $ScriptRelativePath).
        Replace('__LOG_FILE_NAME__', $LogFileName).
        Replace('__HOST_KIND_FOR_STATUS__', $StatusHostKind).
        Replace('__LAUNCH_BLOCK__', $launchBlock)
}

function Write-DistBatLaunchers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistRoot,

        [Parameter(Mandatory = $true)]
        [string]$DistScriptsDir
    )

    Write-TextFileUtf8Bom -Path (Join-Path $DistScriptsDir 'launch_office_host.ps1') -Content (Get-OfficeLaunchScriptContent)
    Write-TextFileUtf8Bom -Path (Join-Path $DistScriptsDir 'show_install_status.ps1') -Content (Get-InstallStatusScriptContent)

    $launchers = @(
        @{
            BatName = 'install_excel_addin.bat'
            ScriptRelativePath = 'files\scripts\install_excel_addin.ps1'
            DisplayName = 'Excel Add-in Install'
            LogFileName = 'install_excel_addin.log'
            StatusHostKind = 'Excel'
            LaunchHostKind = 'Excel'
        },
        @{
            BatName = 'install_wps_addin.bat'
            ScriptRelativePath = 'files\scripts\install_wps_addin.ps1'
            DisplayName = 'WPS Add-in Install'
            LogFileName = 'install_wps_addin.log'
            StatusHostKind = 'WPS'
            LaunchHostKind = 'WPS'
        },
        @{
            BatName = 'uninstall_excel_addin.bat'
            ScriptRelativePath = 'files\scripts\uninstall_excel_addin.ps1'
            DisplayName = 'Excel Add-in Uninstall'
            LogFileName = 'uninstall_excel_addin.log'
            StatusHostKind = 'Excel'
            LaunchHostKind = ''
        },
        @{
            BatName = 'uninstall_wps_addin.bat'
            ScriptRelativePath = 'files\scripts\uninstall_wps_addin.ps1'
            DisplayName = 'WPS Add-in Uninstall'
            LogFileName = 'uninstall_wps_addin.log'
            StatusHostKind = 'WPS'
            LaunchHostKind = ''
        }
    )

    foreach ($launcher in $launchers) {
        $scriptPath = Join-Path $DistRoot $launcher.ScriptRelativePath
        if (-not (Test-Path $scriptPath)) {
            continue
        }

        $content = Get-BatLauncherContent -DisplayName $launcher.DisplayName -ScriptRelativePath $launcher.ScriptRelativePath -LogFileName $launcher.LogFileName -StatusHostKind $launcher.StatusHostKind -LaunchHostKind $launcher.LaunchHostKind
        Write-TextFileAscii -Path (Join-Path $DistRoot $launcher.BatName) -Content $content
    }
}

function Get-ProjectTargetFramework {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectFilePath
    )

    $content = [System.IO.File]::ReadAllText($ProjectFilePath)

    $singleMatch = [System.Text.RegularExpressions.Regex]::Match($content, '<TargetFramework>\s*([^<]+?)\s*</TargetFramework>')
    if ($singleMatch.Success) {
        return $singleMatch.Groups[1].Value.Trim()
    }

    $multiMatch = [System.Text.RegularExpressions.Regex]::Match($content, '<TargetFrameworks>\s*([^<]+?)\s*</TargetFrameworks>')
    if ($multiMatch.Success) {
        $firstFramework = $multiMatch.Groups[1].Value.Split(';') | Select-Object -First 1
        if (-not [string]::IsNullOrWhiteSpace($firstFramework)) {
            return $firstFramework.Trim()
        }
    }

    throw "无法从项目文件中解析 TargetFramework：$ProjectFilePath"
}

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
}

$resolvedProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
$skillToolsRoot = $PSScriptRoot
$syncRuntimeScript = Join-Path $skillToolsRoot 'sync_local_runtime.ps1'
$srcRoot = Join-Path $resolvedProjectRoot 'src'
$projectFile = Get-ChildItem -Path $srcRoot -Recurse -Filter *.csproj | Select-Object -First 1

if ($null -eq $projectFile) {
    throw "未在 src 下找到 csproj：$srcRoot"
}

$targetFramework = Get-ProjectTargetFramework -ProjectFilePath $projectFile.FullName

Assert-DotnetAvailable
& $syncRuntimeScript -ProjectRoot $resolvedProjectRoot -Force

Invoke-DotnetCommand -WorkingDirectory $resolvedProjectRoot -Arguments @('build', $projectFile.FullName, '-c', $Configuration)

$outputDir = Join-Path $projectFile.Directory.FullName ('bin\' + $Configuration + '\' + $targetFramework)
$distDir = Join-Path $resolvedProjectRoot 'dist'
$distFilesDir = Join-Path $distDir 'files'
if (-not (Test-Path $outputDir)) {
    throw "未找到构建输出目录：$outputDir"
}

[void][System.IO.Directory]::CreateDirectory($distDir)
[void][System.IO.Directory]::CreateDirectory($distFilesDir)

$artifacts = @(
    '*.xll',
    '*.dna',
    '*.deps.json',
    '*.runtimeconfig.json',
    'appsettings.json',
    'appsettings.user.json.example',
    '*.dll',
    '*.pdb',
    '*.chm'
)

foreach ($pattern in $artifacts) {
    Get-ChildItem -Path $outputDir -Filter $pattern -File -ErrorAction SilentlyContinue | ForEach-Object {
        $destinationPath = Join-Path $distFilesDir $_.Name
        try {
            [System.IO.File]::Copy($_.FullName, $destinationPath, $true)
        }
        catch [System.IO.IOException] {
            Write-Warning ("目标文件被占用，已跳过复制：{0}" -f $destinationPath)
            Write-Warning '请先关闭正在加载该 XLL 的 Excel/WPS 进程后重新执行 build_addin.ps1。'
        }
    }
}

$runtimeAssetsDir = Join-Path $outputDir 'runtimes'
if (Test-Path $runtimeAssetsDir) {
    $distRuntimeDir = Join-Path $distFilesDir 'runtimes'
    [void][System.IO.Directory]::CreateDirectory($distRuntimeDir)
    Copy-DirectoryArtifacts -SourceRoot $runtimeAssetsDir -TargetRoot $distRuntimeDir
}

$contentAssetsDir = Join-Path $outputDir 'assets'
if (Test-Path $contentAssetsDir) {
    $distAssetsDir = Join-Path $distFilesDir 'assets'
    [void][System.IO.Directory]::CreateDirectory($distAssetsDir)
    Copy-DirectoryArtifacts -SourceRoot $contentAssetsDir -TargetRoot $distAssetsDir
}

$projectScriptsDir = Join-Path $resolvedProjectRoot 'scripts'
$distScriptsDir = Join-Path $distFilesDir 'scripts'
if (Test-Path $projectScriptsDir) {
    [void][System.IO.Directory]::CreateDirectory($distScriptsDir)
    Copy-DirectoryArtifacts -SourceRoot $projectScriptsDir -TargetRoot $distScriptsDir
}

Invoke-ExampleWorkbookGenerator -ProjectRoot $resolvedProjectRoot -ProjectScriptsDir $projectScriptsDir

$projectExamplesDir = Join-Path $resolvedProjectRoot 'examples'
$distExamplesDir = Join-Path $distDir 'examples'
if (Test-Path $projectExamplesDir) {
    [void][System.IO.Directory]::CreateDirectory($distExamplesDir)
    Copy-DirectoryArtifacts -SourceRoot $projectExamplesDir -TargetRoot $distExamplesDir
}

Copy-FileArtifacts -SourceRoot $resolvedProjectRoot -TargetRoot $distDir -Patterns @('*.bat', '安装说明.txt')
Write-DistBatLaunchers -DistRoot $distDir -DistScriptsDir $distScriptsDir

Write-Host "构建完成，交付入口目录：$distDir"
Write-Host "插件文件目录：$distFilesDir"
if (Test-Path $distExamplesDir) {
    Write-Host "示例文件目录：$distExamplesDir"
}

<#
v_id: 125c93a5-7e2dfac4-7c3ba29d-276bcc40-a9ca7619-a803771d-81bb070d
env_hash: 79b577cb-9f29db2f-c4299258-f850ca59-9c30db2f-c5229244-ce9532b3-1ad01b2e-fb199247-ef50fe49-5953fe4b-9f29fe24-c5399370-fc5cee5b-9f3bff2d-e436905f-d153ff7c-9132dd2c-ed1d9877-f552d14a-9f18d52f-c20e9376-ec50ca69-9c09f82c-e3319158-fc5df061-9d09d72d-eb18944b-fb50d872-9d0ff92d-f803914f-f651cb6b-9f27da28-f9349f5b-dc50ff62-9f35d02f-c5159159-d45dd647-9d0dcd24-c5399143-e853e172-9c05f12f-c628905e-e05dc876-9e1cc12d-ca209275-f25dc368-9d0ecc2c-e3319156-fa50ff62-9a35f5
#>
