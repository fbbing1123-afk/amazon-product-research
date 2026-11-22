param(
  [Parameter(Mandatory=$true)][string]$WorkbookPath,
  [Parameter(Mandatory=$true)][string]$OutputDir,
  [string]$ProgId = "Excel.Application"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$app = $null
$workbook = $null
try {
  $app = New-Object -ComObject $ProgId
  $app.Visible = $true
  $app.DisplayAlerts = $false
  $workbook = $app.Workbooks.Open($resolvedWorkbook)

  $sheet = $null
  foreach ($ws in @($workbook.Worksheets)) {
    try {
      if ($ws.ChartObjects().Count -gt 0) {
        $sheet = $ws
        break
      }
    } catch {
    }
  }
  if ($null -eq $sheet) {
    throw "No sheet with chart objects was found."
  }

  $sheet.Activate() | Out-Null
  $count = $sheet.ChartObjects().Count
  for ($i = 1; $i -le $count; $i++) {
    $chartObject = $sheet.ChartObjects($i)
    $chartObject.Activate() | Out-Null
    [System.Windows.Forms.Clipboard]::Clear()
    $chartObject.CopyPicture(1, 2) | Out-Null
    Start-Sleep -Milliseconds 500
    $image = [System.Windows.Forms.Clipboard]::GetImage()
    if ($null -eq $image) {
      throw "Clipboard did not contain an image for chart $i."
    }
    $path = Join-Path $OutputDir ("market_chart_{0}.png" -f $i)
    $image.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $image.Dispose()
    Write-Host $path
  }
}
finally {
  if ($workbook -ne $null) { $workbook.Close($false) | Out-Null }
  if ($app -ne $null) { $app.Quit() | Out-Null }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
