#requires -Version 5.1

$ErrorActionPreference = 'Stop'

function Get-RegistryViews {
    if ([Environment]::Is64BitOperatingSystem) {
        return @(
            [Microsoft.Win32.RegistryView]::Registry64,
            [Microsoft.Win32.RegistryView]::Registry32
        )
    }

    return @([Microsoft.Win32.RegistryView]::Default)
}

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

function Set-RegistryStringValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$SubKey,
        [string]$Name,
        [string]$Value
    )

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
    try {
        $key = $baseKey.CreateSubKey($SubKey)
        try {
            $key.SetValue($Name, $Value, [Microsoft.Win32.RegistryValueKind]::String)
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

function Remove-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [Microsoft.Win32.RegistryView]$View,
        [string]$SubKey,
        [string]$Name
    )

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
    try {
        $key = $baseKey.OpenSubKey($SubKey, $true)
        try {
            if ($null -ne $key -and $key.GetValue($Name, $null) -ne $null) {
                $key.DeleteValue($Name, $false)
            }
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

function Normalize-OfficeArchitecture {
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    $normalized = ([string]$Value).Trim().ToLowerInvariant()
    switch ($normalized) {
        'x64' { return 'x64' }
        'amd64' { return 'x64' }
        '64' { return 'x64' }
        '64bit' { return 'x64' }
        '64-bit' { return 'x64' }
        'x86' { return 'x86' }
        '32' { return 'x86' }
        '32bit' { return 'x86' }
        '32-bit' { return 'x86' }
        default { return '' }
    }
}

function Get-RegistryViewForArchitecture {
    param(
        [ValidateSet('x86', 'x64')]
        [string]$Architecture
    )

    if ($Architecture -eq 'x64') {
        if (-not [Environment]::Is64BitOperatingSystem) {
            throw '当前系统不是 64 位，不能安装 x64 版本。'
        }

        return [Microsoft.Win32.RegistryView]::Registry64
    }

    if ([Environment]::Is64BitOperatingSystem) {
        return [Microsoft.Win32.RegistryView]::Registry32
    }

    return [Microsoft.Win32.RegistryView]::Default
}

function Get-DistRoot {
    param(
        [string]$ProjectName
    )

    $scriptParent = Split-Path $PSScriptRoot -Parent
    $nestedDist = Join-Path $scriptParent 'dist'
    $nestedFilesDist = Join-Path $nestedDist 'files'
    $directX86 = Join-Path $scriptParent ($ProjectName + '-AddIn.xll')
    $directX64 = Join-Path $scriptParent ($ProjectName + '-AddIn64.xll')
    $nestedX86 = Join-Path $nestedDist ($ProjectName + '-AddIn.xll')
    $nestedX64 = Join-Path $nestedDist ($ProjectName + '-AddIn64.xll')
    $nestedFilesX86 = Join-Path $nestedFilesDist ($ProjectName + '-AddIn.xll')
    $nestedFilesX64 = Join-Path $nestedFilesDist ($ProjectName + '-AddIn64.xll')

    if ((Test-Path $directX86) -or (Test-Path $directX64)) {
        return $scriptParent
    }

    if ((Test-Path $nestedX86) -or (Test-Path $nestedX64)) {
        return $nestedDist
    }

    if ((Test-Path $nestedFilesX86) -or (Test-Path $nestedFilesX64)) {
        return $nestedFilesDist
    }

    throw "未找到 dist 产物目录，请先运行 build_addin.ps1。已检查: $scriptParent、$nestedDist 和 $nestedFilesDist"
}

function Get-AddinArtifacts {
    param(
        [string]$ProjectName
    )

    $distRoot = Get-DistRoot -ProjectName $ProjectName
    $x86Path = Join-Path $distRoot ($ProjectName + '-AddIn.xll')
    $x64Path = Join-Path $distRoot ($ProjectName + '-AddIn64.xll')

    return [pscustomobject]@{
        DistRoot = $distRoot
        X86 = [pscustomobject]@{
            Architecture = 'x86'
            XllPath = $x86Path
            Exists = (Test-Path $x86Path)
        }
        X64 = [pscustomobject]@{
            Architecture = 'x64'
            XllPath = $x64Path
            Exists = (Test-Path $x64Path)
        }
    }
}

function Get-AddinSupportPaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistRoot
    )

    return [pscustomobject]@{
        LogsRoot = Join-Path $DistRoot 'logs'
        WebView2Root = Join-Path $DistRoot 'webview2'
    }
}

function Remove-WebView2UserData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DistRoot
    )

    $supportPaths = Get-AddinSupportPaths -DistRoot $DistRoot
    if (Test-Path -LiteralPath $supportPaths.WebView2Root) {
        [System.IO.Directory]::Delete($supportPaths.WebView2Root, $true)
    }
}

function Resolve-AddinPathForArchitecture {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Artifacts,
        [ValidateSet('x86', 'x64')]
        [string]$Architecture
    )

    $candidate = if ($Architecture -eq 'x64') { $Artifacts.X64 } else { $Artifacts.X86 }
    if (-not $candidate.Exists) {
        throw "未找到 $Architecture 版本的 XLL：$($candidate.XllPath)"
    }

    return [System.IO.Path]::GetFullPath($candidate.XllPath)
}

function Get-ExecutableArchitecture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return ''
    }

    $stream = $null
    $reader = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.BinaryReader($stream)
        $stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peOffset = $reader.ReadInt32()
        $stream.Seek($peOffset + 4, [System.IO.SeekOrigin]::Begin) | Out-Null
        $machine = $reader.ReadUInt16()

        switch ($machine) {
            0x014c { return 'x86' }
            0x8664 { return 'x64' }
            default { return '' }
        }
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Start-TrackedComApplication {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgId,
        [Parameter(Mandatory = $true)]
        [string]$ProcessName
    )

    $beforeIds = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $application = New-Object -ComObject $ProgId
    $processId = $null
    $processPath = ''
    $deadline = (Get-Date).AddSeconds(10)

    while ($null -eq $processId -and (Get-Date) -lt $deadline) {
        $candidates = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $beforeIds -notcontains $_.Id })
        if ($candidates.Count -gt 0) {
            $process = $candidates | Sort-Object StartTime -Descending | Select-Object -First 1
            $processId = $process.Id
            try {
                $processPath = [string]$process.Path
            }
            catch {
                $processPath = ''
            }
        }
        else {
            Start-Sleep -Milliseconds 200
        }
    }

    return [pscustomobject]@{
        Application = $application
        ProcessId = $processId
        ProcessPath = $processPath
        ProgId = $ProgId
    }
}

function Stop-TrackedComApplication {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TrackedApp
    )

    if ($null -ne $TrackedApp.Application) {
        try {
            $TrackedApp.Application.Quit()
        }
        catch {
        }
        finally {
            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($TrackedApp.Application)
            }
            catch {
            }
        }
    }

    if ($TrackedApp.ProcessId) {
        try {
            Stop-Process -Id $TrackedApp.ProcessId -Force -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 300
}

function Try-Get-ExcelBitnessFromVbaRun {
    $trackedExcel = $null
    $workbook = $null
    $vbProject = $null
    $module = $null

    try {
        $trackedExcel = Start-TrackedComApplication -ProgId 'Excel.Application' -ProcessName 'EXCEL'
        $excel = $trackedExcel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $vbProject = $workbook.VBProject
        $module = $vbProject.VBComponents.Add(1)
        $module.Name = 'XlDnaBitnessProbe'
        $vbCode = @(
            'Public Function __XlDnaDetectOfficeBitness() As String',
            '#If Win64 Then',
            '    __XlDnaDetectOfficeBitness = "x64"',
            '#Else',
            '    __XlDnaDetectOfficeBitness = "x86"',
            '#End If',
            'End Function'
        ) -join [Environment]::NewLine
        $module.CodeModule.AddFromString($vbCode)

        $macroName = "'" + [string]$workbook.Name + "'!__XlDnaDetectOfficeBitness"
        $result = Normalize-OfficeArchitecture -Value ([string]$excel.Run($macroName))
        if (-not [string]::IsNullOrWhiteSpace($result)) {
            return [pscustomobject]@{
                Architecture = $result
                Source = 'COM.Application.Run'
                Evidence = '通过临时 VBA 函数 __XlDnaDetectOfficeBitness 返回'
            }
        }
    }
    catch {
    }
    finally {
        if ($null -ne $workbook) {
            try {
                $workbook.Close($false)
            }
            catch {
            }
        }

        foreach ($comObject in @($module, $vbProject, $workbook)) {
            if ($null -ne $comObject) {
                try {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
                }
                catch {
                }
            }
        }

        if ($null -ne $trackedExcel) {
            Stop-TrackedComApplication -TrackedApp $trackedExcel
        }
    }

    return $null
}

function Get-ExcelBitnessFromClickToRun {
    $checks = @(
        @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry64; SubKey = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'Platform' },
        @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry32; SubKey = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Name = 'Platform' },
        @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry64; SubKey = 'SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'; Name = 'Platform' }
    )

    foreach ($check in $checks) {
        if (-not [Environment]::Is64BitOperatingSystem -and $check.View -eq [Microsoft.Win32.RegistryView]::Registry64) {
            continue
        }

        $platform = Normalize-OfficeArchitecture -Value (Get-RegistryStringValue -Hive $check.Hive -View $check.View -SubKey $check.SubKey -Name $check.Name)
        if (-not [string]::IsNullOrWhiteSpace($platform)) {
            return [pscustomobject]@{
                Architecture = $platform
                Source = 'Registry.ClickToRun'
                Evidence = $check.SubKey + '\' + $check.Name
            }
        }
    }

    return $null
}

function Get-ExcelBitnessFromExecutablePath {
    $candidates = @()
    $candidates += Get-RegistryStringValue -Hive LocalMachine -View Registry64 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -Name ''
    $candidates += Get-RegistryStringValue -Hive LocalMachine -View Registry32 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -Name ''
    $candidates += Get-RegistryStringValue -Hive CurrentUser -View Registry64 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -Name ''
    $candidates += Get-RegistryStringValue -Hive CurrentUser -View Registry32 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -Name ''

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        $architecture = Get-ExecutableArchitecture -Path $candidate
        if (-not [string]::IsNullOrWhiteSpace($architecture)) {
            return [pscustomobject]@{
                Architecture = $architecture
                Source = 'Registry.AppPaths'
                Evidence = $candidate
            }
        }
    }

    return $null
}

function Get-ExcelBitnessFromLiveProcess {
    $trackedExcel = $null

    try {
        $trackedExcel = Start-TrackedComApplication -ProgId 'Excel.Application' -ProcessName 'EXCEL'
        $processPath = $trackedExcel.ProcessPath
        if ([string]::IsNullOrWhiteSpace($processPath) -and $trackedExcel.ProcessId) {
            try {
                $processPath = [string](Get-Process -Id $trackedExcel.ProcessId -ErrorAction Stop).Path
            }
            catch {
                $processPath = ''
            }
        }

        $architecture = Get-ExecutableArchitecture -Path $processPath
        if (-not [string]::IsNullOrWhiteSpace($architecture)) {
            return [pscustomobject]@{
                Architecture = $architecture
                Source = 'Process.EXCEL'
                Evidence = $processPath
            }
        }
    }
    catch {
    }
    finally {
        if ($null -ne $trackedExcel) {
            Stop-TrackedComApplication -TrackedApp $trackedExcel
        }
    }

    return $null
}

function Get-ExcelBitness {
    $detectors = @(
        { Get-ExcelBitnessFromClickToRun },
        { Get-ExcelBitnessFromExecutablePath },
        { Try-Get-ExcelBitnessFromVbaRun },
        { Get-ExcelBitnessFromLiveProcess }
    )

    foreach ($detector in $detectors) {
        $result = & $detector
        if ($null -ne $result) {
            return $result
        }
    }

    throw '无法自动判断 Excel 是 32 位还是 64 位，请手动传入 -Architecture x86 或 -Architecture x64。'
}

function Get-WpsBitness {
    foreach ($view in (Get-RegistryViews)) {
        $architecture = Normalize-OfficeArchitecture -Value (Get-RegistryStringValue -Hive CurrentUser -View $view -SubKey 'Software\kingsoft\office\6.0\Common' -Name 'Architecture')
        if (-not [string]::IsNullOrWhiteSpace($architecture)) {
            return [pscustomobject]@{
                Architecture = $architecture
                Source = 'Registry.WPS.Common'
                Evidence = 'Software\kingsoft\office\6.0\Common\Architecture'
            }
        }
    }

    $candidates = @()
    $candidates += Get-RegistryStringValue -Hive CurrentUser -View Registry64 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\et.exe' -Name ''
    $candidates += Get-RegistryStringValue -Hive CurrentUser -View Registry32 -SubKey 'Software\Microsoft\Windows\CurrentVersion\App Paths\et.exe' -Name ''
    $candidates += Get-RegistryStringValue -Hive LocalMachine -View Registry64 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\et.exe' -Name ''
    $candidates += Get-RegistryStringValue -Hive LocalMachine -View Registry32 -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\et.exe' -Name ''

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        $architecture = Get-ExecutableArchitecture -Path $candidate
        if (-not [string]::IsNullOrWhiteSpace($architecture)) {
            return [pscustomobject]@{
                Architecture = $architecture
                Source = 'Executable.ET'
                Evidence = $candidate
            }
        }
    }

    throw '无法自动判断 WPS ET 是 32 位还是 64 位，请手动传入 -Architecture x86 或 -Architecture x64。'
}

function Resolve-TargetArchitecture {
    param(
        [ValidateSet('Auto', 'x86', 'x64')]
        [string]$Architecture,
        [Parameter(Mandatory = $true)]
        [object]$DetectedInfo
    )

    if ($Architecture -eq 'Auto') {
        return $DetectedInfo.Architecture
    }

    return $Architecture
}

function Get-ExcelRegistryVersions {
    return @('16.0', '15.0', '14.0', '12.0')
}

function Ensure-ExcelOpenEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OptionsPath,
        [Parameter(Mandatory = $true)]
        [string]$AddinPath
    )

    $addinName = [System.IO.Path]::GetFileName($AddinPath)
    $current = Get-ItemProperty -LiteralPath $OptionsPath -ErrorAction SilentlyContinue
    $openNames = @()

    if ($null -ne $current) {
        $openNames = @(
            $current.PSObject.Properties |
            Where-Object { $_.Name -match '^OPEN\d*$' } |
            Sort-Object {
                if ($_.Name -eq 'OPEN') { 0 } else { [int]$_.Name.Substring(4) }
            }
        )
    }

    foreach ($property in $openNames) {
        $valueText = [string]$property.Value
        if ($valueText -like ('*' + $AddinPath + '*') -or $valueText -like ('*' + $addinName + '*')) {
            return
        }
    }

    $slotName = 'OPEN'
    $index = 0
    while ($null -ne (Get-ItemProperty -LiteralPath $OptionsPath -Name $slotName -ErrorAction SilentlyContinue)) {
        $index++
        $slotName = 'OPEN' + $index
    }

    New-ItemProperty -LiteralPath $OptionsPath -Name $slotName -Value ('/R "' + $AddinPath + '"') -PropertyType String -Force | Out-Null
}

function Register-ExcelAddinRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AddinPath
    )

    foreach ($version in (Get-ExcelRegistryVersions)) {
        $optionsPath = 'HKCU:\Software\Microsoft\Office\' + $version + '\Excel\Options'
        $managerPath = 'HKCU:\Software\Microsoft\Office\' + $version + '\Excel\Add-in Manager'

        if (-not (Test-Path -LiteralPath $optionsPath)) {
            New-Item -Path $optionsPath -Force | Out-Null
        }
        if (-not (Test-Path -LiteralPath $managerPath)) {
            New-Item -Path $managerPath -Force | Out-Null
        }

        New-ItemProperty -LiteralPath $managerPath -Name $AddinPath -Value '' -PropertyType String -Force | Out-Null
        Ensure-ExcelOpenEntry -OptionsPath $optionsPath -AddinPath $AddinPath
    }
}

function Unregister-ExcelAddinRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$AddinPaths
    )

    $addins = @($AddinPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($addins.Count -eq 0) {
        return
    }

    $addinNames = @($addins | ForEach-Object { [System.IO.Path]::GetFileName($_) })

    foreach ($version in (Get-ExcelRegistryVersions)) {
        $optionsPath = 'HKCU:\Software\Microsoft\Office\' + $version + '\Excel\Options'
        $managerPath = 'HKCU:\Software\Microsoft\Office\' + $version + '\Excel\Add-in Manager'

        if (Test-Path -LiteralPath $optionsPath) {
            $current = Get-ItemProperty -LiteralPath $optionsPath
            foreach ($property in @($current.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })) {
                $valueText = [string]$property.Value
                if (@($addins | Where-Object { $valueText -like ('*' + $_ + '*') }).Count -gt 0) {
                    Remove-ItemProperty -LiteralPath $optionsPath -Name $property.Name -ErrorAction SilentlyContinue
                    continue
                }
                if (@($addinNames | Where-Object { $valueText -like ('*' + $_ + '*') }).Count -gt 0) {
                    Remove-ItemProperty -LiteralPath $optionsPath -Name $property.Name -ErrorAction SilentlyContinue
                }
            }
        }

        if (Test-Path -LiteralPath $managerPath) {
            $current = Get-ItemProperty -LiteralPath $managerPath
            foreach ($property in @($current.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' })) {
                if ($property.Name -in $addins) {
                    Remove-ItemProperty -LiteralPath $managerPath -Name $property.Name -ErrorAction SilentlyContinue
                    continue
                }
                if (@($addinNames | Where-Object { $property.Name -like ('*' + $_) }).Count -gt 0) {
                    Remove-ItemProperty -LiteralPath $managerPath -Name $property.Name -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

function Register-WpsAddin {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AddinPath,
        [ValidateSet('x86', 'x64')]
        [string]$Architecture,
        [string[]]$ProgIds = @()
    )

    $view = Get-RegistryViewForArchitecture -Architecture $Architecture
    Set-RegistryStringValue -Hive CurrentUser -View $view -SubKey 'Software\kingsoft\office\6.0\et\LoadMacros' -Name $AddinPath -Value '1'

    foreach ($progId in @($ProgIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        Set-RegistryStringValue -Hive CurrentUser -View $view -SubKey 'Software\kingsoft\office\ET\AddinsWL' -Name $progId -Value $progId
    }
}

function Unregister-WpsAddin {
    param(
        [string[]]$AddinPaths = @(),
        [string[]]$ProgIds = @()
    )

    foreach ($view in (Get-RegistryViews)) {
        foreach ($addinPath in @($AddinPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            Remove-RegistryValue -Hive CurrentUser -View $view -SubKey 'Software\kingsoft\office\6.0\et\LoadMacros' -Name $addinPath
        }
        foreach ($progId in @($ProgIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            Remove-RegistryValue -Hive CurrentUser -View $view -SubKey 'Software\kingsoft\office\ET\AddinsWL' -Name $progId
        }
    }
}

function Get-ProjectDataDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectName
    )

    $distRoot = Get-DistRoot -ProjectName $ProjectName
    return [pscustomobject]@{
        DistRoot = $distRoot
        LogsRoot = Join-Path $distRoot 'logs'
        WebView2Root = Join-Path $distRoot 'webview2'
    }
}

function Remove-ProjectGeneratedData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectName
    )

    $dataDirectories = Get-ProjectDataDirectories -ProjectName $ProjectName
    foreach ($path in @($dataDirectories.LogsRoot, $dataDirectories.WebView2Root)) {
        if (Test-Path -LiteralPath $path) {
            try {
                [System.IO.Directory]::Delete($path, $true)
                Write-Host ("已清理目录：{0}" -f $path)
            }
            catch {
                Write-Warning ("清理目录失败：{0}，原因：{1}" -f $path, $_.Exception.Message)
            }
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

function Start-OfficeHost {
    param(
        [ValidateSet('Excel', 'WPS')]
        [string]$HostKind
    )

    $processPath = if ($HostKind -eq 'Excel') { Get-ExcelExecutablePath } else { Get-WpsExecutablePath }
    if ([string]::IsNullOrWhiteSpace($processPath)) {
        return [pscustomobject]@{
            Started = $false
            Path = ''
            Message = "未找到 $HostKind 可执行文件，请手动启动。"
        }
    }

    try {
        Start-Process -FilePath $processPath | Out-Null
        return [pscustomobject]@{
            Started = $true
            Path = $processPath
            Message = "已尝试启动 $HostKind。"
        }
    }
    catch {
        return [pscustomobject]@{
            Started = $false
            Path = $processPath
            Message = ("启动 {0} 失败：{1}" -f $HostKind, $_.Exception.Message)
        }
    }
}

function Write-InstallSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostKind,
        [Parameter(Mandatory = $true)]
        [string]$Architecture,
        [Parameter(Mandatory = $true)]
        [string]$Source,
        [Parameter(Mandatory = $true)]
        [string]$Evidence,
        [Parameter(Mandatory = $true)]
        [string]$XllPath,
        [Parameter(Mandatory = $true)]
        [string]$ActionSummary
    )

    Write-Host ''
    Write-Host '========================================'
    Write-Host ($HostKind + ' 安装完成')
    Write-Host '========================================'
    Write-Host ('宿主位数：{0}' -f $Architecture)
    Write-Host ('检测来源：{0}' -f $Source)
    Write-Host ('检测证据：{0}' -f $Evidence)
    Write-Host ('加载项路径：{0}' -f $XllPath)
    Write-Host ('安装动作：{0}' -f $ActionSummary)
}
