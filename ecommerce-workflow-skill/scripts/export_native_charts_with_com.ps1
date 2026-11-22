param(
  [Parameter(Mandatory=$true)][string]$WorkbookPath,
  [Parameter(Mandatory=$true)][string]$OutputDir,
  [string]$SheetSuffix = "市场分析",
  [string]$ProgId = "Excel.Application"
)

$ErrorActionPreference = "Stop"
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$stagingDir = Join-Path $env:ProgramData "CodexChartExports"
New-Item -ItemType Directory -Force -Path $stagingDir | Out-Null

$app = $null
$workbook = $null
try {
  $app = New-Object -ComObject $ProgId
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $workbook = $app.Workbooks.Open($resolvedWorkbook)
  $sheet = $null
  foreach ($ws in @($workbook.Worksheets)) {
    if ([string]$ws.Name -like "*$SheetSuffix") {
      $sheet = $ws
      break
    }
  }
  if ($null -eq $sheet) {
    foreach ($ws in @($workbook.Worksheets)) {
      try {
        if ($ws.ChartObjects().Count -gt 0) {
          $sheet = $ws
          break
        }
      } catch {
      }
    }
  }
  if ($null -eq $sheet) {
    throw "No sheet with chart objects was found."
  }

  $count = $sheet.ChartObjects().Count
  if ($count -lt 1) {
    throw "No chart objects found on sheet $($sheet.Name)."
  }

  for ($i = 1; $i -le $count; $i++) {
    $chartObject = $sheet.ChartObjects($i)
    $path = Join-Path $OutputDir ("market_chart_{0}.png" -f $i)
    $stagingPath = Join-Path $stagingDir ("market_chart_{0}.png" -f $i)
    if (Test-Path -LiteralPath $stagingPath) {
      Remove-Item -LiteralPath $stagingPath -Force
    }
    $chartObject.Activate() | Out-Null
    $ok = $chartObject.Chart.Export($stagingPath, "PNG", $false)
    if (-not $ok) {
      throw "Failed to export chart $i to $stagingPath"
    }
    Copy-Item -LiteralPath $stagingPath -Destination $path -Force
    Write-Host $path
  }
}
finally {
  if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }
  if ($app -ne $null) { $app.Quit() | Out-Null }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
