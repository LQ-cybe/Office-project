param(
    [string]$ProjectRoot = (Join-Path $PSScriptRoot '..'),
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [string]$TargetFramework = 'net8.0-windows',
    [switch]$SkipBuild,
    [switch]$Visible
)

# 说明：
# 1. 这是蓝图示例脚本，当前默认覆盖 UDF、命令型入口、结构化表输出与部分对象模型操作。
# 2. 正式项目使用前，必须先根据当前项目真实导出的函数名、Smoke 命令名、工作表路径和保留模块重建测试内容。
# 3. 如果项目删掉了图表、CTP、WebView2 或其他默认示例模块，请同步删改对应断言。
# 4. 如果项目新增了 RTD、动态数组或其他自定义能力，建议补充新的公式验证，而不是继续依赖模板默认假设。

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-NormalizedPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return [System.IO.Path]::GetFullPath($Path)
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

if (-not ('ExcelSmokeNativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelSmokeNativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
}
'@
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

function Get-ExcelBitnessTag {
    param([Parameter(Mandatory = $true)][string]$ExcelInstallPath)

    if ($ExcelInstallPath.IndexOf('(x86)', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        return 'x86'
    }

    return 'x64'
}

function Convert-RgbToOleColor {
    param(
        [Parameter(Mandatory = $true)][int]$Red,
        [Parameter(Mandatory = $true)][int]$Green,
        [Parameter(Mandatory = $true)][int]$Blue
    )

    return ($Blue * 65536) + ($Green * 256) + $Red
}

function Get-OfficeProcessIds {
    $processes = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match '^(EXCEL|ET|WPS)$' }
    return @($processes | ForEach-Object { [int]$_.Id })
}

function Get-ExcelApplicationProcessId {
    param([Parameter(Mandatory = $true)][object]$ExcelApplication)

    try {
        $hwnd = [System.IntPtr]::new([int]$ExcelApplication.Hwnd)
        if ($hwnd -eq [System.IntPtr]::Zero) {
            return $null
        }

        $processId = 0
        [void][ExcelSmokeNativeMethods]::GetWindowThreadProcessId($hwnd, [ref]$processId)
        if ($processId -gt 0) {
            return [int]$processId
        }
    }
    catch {
    }

    return $null
}

function Start-NewExcelApplication {
    param([switch]$VisibleWindow)

    $beforeIds = @(Get-OfficeProcessIds)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$VisibleWindow
    $excel.DisplayAlerts = $false

    Start-Sleep -Milliseconds 500

    $afterIds = @(Get-OfficeProcessIds)
    $trackedIds = @($afterIds | Where-Object { $beforeIds -notcontains $_ })
    $excelProcessId = Get-ExcelApplicationProcessId -ExcelApplication $excel

    if ($excelProcessId -and $beforeIds -notcontains $excelProcessId) {
        $trackedIds = @($excelProcessId) + @($trackedIds | Where-Object { $_ -ne $excelProcessId })
    }

    return @{
        Excel = $excel
        TrackedProcessIds = @($trackedIds | Select-Object -Unique)
    }
}

function Stop-TrackedOfficeProcesses {
    param([int[]]$ProcessIds)

    foreach ($processId in @($ProcessIds | Where-Object { $_ -gt 0 } | Select-Object -Unique)) {
        try {
            $process = Get-Process -Id $processId -ErrorAction Stop
            if (-not $process.HasExited) {
                Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
        }
    }
}

function Get-XllPath {
    param(
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $true)][string]$BitnessTag
    )

    $fileName = if ($BitnessTag -eq 'x86') { 'MultiTableAddin-AddIn.xll' } else { 'MultiTableAddin-AddIn64.xll' }
    $candidate = Join-Path $OutputDirectory $fileName
    if (-not (Test-Path -LiteralPath $candidate)) {
        throw "未找到待测 XLL：$candidate"
    }

    return $candidate
}

$projectRootPath = Resolve-NormalizedPath -Path $ProjectRoot
$projectFile = Join-Path $projectRootPath 'src\MultiTableAddin\MultiTableAddin.csproj'
$outputDirectory = Join-Path $projectRootPath ('src\MultiTableAddin\bin\' + $Configuration + '\' + $TargetFramework)

if (-not (Test-Path -LiteralPath $projectFile)) {
    throw "未找到项目文件：$projectFile"
}

if (-not $SkipBuild) {
    if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
        throw '当前机器未找到 dotnet，请先执行 ensure_dotnet_sdk.ps1。'
    }

    Write-Host '开始构建 Debug/Smoke 产物...'
    & dotnet build $projectFile -c $Configuration /p:EnableSmokeTestHooks=true
    if ($LASTEXITCODE -ne 0) {
        throw 'dotnet build 失败，已停止冒烟测试。'
    }
}

$excel = $null
$workbook = $null
$sheet = $null
$exportSheet = $null
$exportTable = $null
$rewriteTable = $null
$chartObject = $null
$chart = $null
$trackedProcessIds = @()

try {
    $excelSession = Start-NewExcelApplication -VisibleWindow:$Visible
    $excel = $excelSession.Excel
    $trackedProcessIds = @($excelSession.TrackedProcessIds)

    $bitnessTag = Get-ExcelBitnessTag -ExcelInstallPath ([string]$excel.Path)
    $xllPath = Get-XllPath -OutputDirectory $outputDirectory -BitnessTag $bitnessTag

    $loaded = $excel.RegisterXLL($xllPath)
    Assert-Condition -Condition ([bool]$loaded) -Message ("XLL 加载失败：{0}" -f $xllPath)

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Smoke'

    $sheet.Range('A1').Formula = '=MULTITABLEADDIN_PARAMS_SUM(2,3,4)'
    $sheet.Range('A2').Formula = '=MULTITABLEADDIN_HELLO("AI")'
    $sheet.Range('A3').Formula = '=MULTITABLEADDIN_LOGPATH()'

    $excel.CalculateFullRebuild()
    Start-Sleep -Seconds 1

    Assert-Condition -Condition ([double]$sheet.Range('A1').Value2 -eq 9) -Message 'PARAMS_SUM 结果不等于 9。'
    Assert-Condition -Condition ((([string]$sheet.Range('A2').Value2) -like '*AI*')) -Message 'HELLO("AI") 未返回预期文本。'
    $logPath = [string]$sheet.Range('A3').Value2
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($logPath)) -Message 'LOGPATH 返回空字符串。'
    $logDirectory = [System.IO.Path]::GetDirectoryName($logPath)
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($logDirectory)) -Message 'LOGPATH 未返回有效目录。'
    Assert-Condition -Condition ([System.IO.Directory]::Exists($logDirectory)) -Message 'LOGPATH 指向的日志目录不存在。'

    $excel.Run('MULTITABLEADDIN_SMOKE_EXPORT_SAMPLE_TABLE')
    Start-Sleep -Milliseconds 500

    $exportSheet = $workbook.Worksheets.Item('自动化结果')
    $exportTable = $exportSheet.ListObjects.Item('tblWorkbookOrders')

    Assert-Condition -Condition ($null -ne $exportTable) -Message '未生成 tblWorkbookOrders。'
    Assert-Condition -Condition ([string]$exportSheet.Range('A2').Text -eq '000000000000000123') -Message '导出表格未按文本列保留长编号。'
    Assert-Condition -Condition ([string]$exportSheet.Range('A2').NumberFormat -eq '@') -Message '订单号列未设置文本格式。'
    Assert-Condition -Condition ([string]$exportSheet.Range('C2').NumberFormat -eq 'yyyy-mm-dd') -Message '业务日期列未设置日期格式。'
    Assert-Condition -Condition ([string]$exportSheet.Range('D2').NumberFormat -eq 'yyyy-mm-dd hh:mm:ss') -Message '更新时间列未设置日期时间格式。'
    Assert-Condition -Condition ([string]$exportSheet.Range('F2').NumberFormat -eq 'hh:mm:ss') -Message '完成时间列未设置时间格式。'

    $sheet.Activate() | Out-Null
    $excel.Run('MULTITABLEADDIN_SMOKE_PREPARE_REWRITE_TABLE')
    $excel.Run('MULTITABLEADDIN_SMOKE_REWRITE_PREPARED_TABLE')
    Start-Sleep -Milliseconds 500

    $rewriteTable = $sheet.ListObjects.Item('tblWorkbookRewrite')
    Assert-Condition -Condition ($null -ne $rewriteTable) -Message '未生成 tblWorkbookRewrite。'
    Assert-Condition -Condition ([string]$sheet.Range('J2').Text -eq '000000000000000123') -Message '重写后的订单号未保留为文本。'
    Assert-Condition -Condition ([string]$sheet.Range('J2').NumberFormat -eq '@') -Message '重写表订单号列未保持文本格式。'
    Assert-Condition -Condition ((([string]$sheet.Range('L2').FormulaR1C1) -like '=IF*')) -Message '金额校验公式列未恢复。'
    Assert-Condition -Condition ((([string]$sheet.Range('O2').FormulaR1C1) -like '=TEXT*')) -Message '状态说明公式列未恢复。'
    Assert-Condition -Condition ([string]$sheet.Range('N2').NumberFormat -eq '#,##0.00') -Message '金额列格式未保留。'
    Assert-Condition -Condition ([string]$sheet.Range('M2').NumberFormat -eq 'yyyy-mm-dd') -Message '重写表业务日期列未设置日期格式。'
    Assert-Condition -Condition ([string]$sheet.Range('P2').NumberFormat -eq 'yyyy-mm-dd hh:mm:ss') -Message '重写表完成时间列未设置日期时间格式。'

    Write-Host 'Excel 冒烟测试通过'
    Write-Host ('宿主位数：{0}' -f $bitnessTag)
    Write-Host ('XLL 路径：{0}' -f $xllPath)
    Write-Host ('日志目录：{0}' -f $logDirectory)
    Write-Host '说明：已覆盖结构化表导出与 ListObject 重写，Ribbon 排版、WebView2 生命周期、CTP 展示仍建议人工验收。'
}
finally {
    if ($workbook -ne $null) {
        try {
            $workbook.Close($false)
        }
        catch {
        }
    }

    Release-ComObject -ComObject $rewriteTable
    Release-ComObject -ComObject $exportTable
    Release-ComObject -ComObject $exportSheet
    Release-ComObject -ComObject $chart
    Release-ComObject -ComObject $chartObject
    Release-ComObject -ComObject $sheet
    Release-ComObject -ComObject $workbook

    if ($excel -ne $null) {
        try {
            $excel.Quit()
        }
        catch {
        }
    }

    Release-ComObject -ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    Stop-TrackedOfficeProcesses -ProcessIds $trackedProcessIds
}
