param([Parameter(Mandatory=$true)][string]$TestRoot)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'browser_runtime.ps1')
. (Join-Path $PSScriptRoot 'runtime_environment.ps1')

function Assert-BrowserTest {
  param([bool]$Passed, [string]$Message)
  if (-not $Passed) { throw $Message }
}

$browserExecutable = Find-WorkflowBrowser
if (-not $browserExecutable) { throw '实测需要本机Chrome或Edge。' }
if (Test-Path -LiteralPath $TestRoot) { throw '测试目录已存在，不覆盖。' }
$TestRoot = [System.IO.Path]::GetFullPath($TestRoot)
$testScripts = Join-Path $TestRoot 'skill\scripts'
$testProfile = Join-Path $TestRoot 'profile'
$testConfig = Join-Path $TestRoot 'skill\.workflow-state\sellersprite\browser-config.json'
New-Item -ItemType Directory -Force -Path $testScripts, $testProfile, (Split-Path -Parent $testConfig) | Out-Null
foreach ($name in @('open_sellersprite_login_browser.ps1', 'browser_runtime.ps1')) {
  Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $testScripts $name)
}
@{ profile_dir=$testProfile; marker='保留旧配置字段' } | ConvertTo-Json | Set-Content -LiteralPath $testConfig -Encoding UTF8
$global:WorkflowBrowserTestRoot = $TestRoot
$global:WorkflowBrowserTestProcesses = New-Object 'System.Collections.Generic.List[int]'

# 测试只启动独立的无界面浏览器，不操作用户原有项目窗口。
function Start-Process {
  param([string]$FilePath, [string[]]$ArgumentList)
  $arguments = @('--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-background-networking') + @($ArgumentList | Where-Object { $_ -ne 'about:blank' -and $_ -ne '--new-window' }) + @('about:blank')
  $global:WorkflowBrowserTestArguments = $arguments
  $child = Microsoft.PowerShell.Management\Start-Process -FilePath $FilePath -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardError (Join-Path $global:WorkflowBrowserTestRoot 'browser.stderr.log')
  $global:WorkflowBrowserTestProcesses.Add($child.Id)
}

$reserved = New-Object 'System.Collections.Generic.List[System.Net.Sockets.TcpListener]'
try {
  # 模拟常用备用端口全部占用，已有其它服务保持不动。
  foreach ($candidate in 9223..9243) {
    $reservation = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $candidate)
    try {
      $reservation.Server.ExclusiveAddressUse = $true
      $reservation.Start()
      $reserved.Add($reservation)
    } catch { $reservation.Stop() }
  }
  $systemPort = Find-FreeBrowserPort
  Assert-BrowserTest ($systemPort -lt 9223 -or $systemPort -gt 9243) '备用范围耗尽后未使用系统分配端口'
  Assert-BrowserTest (Test-BrowserPortAvailable $systemPort) '系统分配的端口不可用'
} finally {
  foreach ($reservation in $reserved) { $reservation.Stop() }
}

$occupied = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 9223)
try {
  $occupied.Server.ExclusiveAddressUse = $true
  try { $occupied.Start() } catch {
    $occupied.Stop()
    $occupied = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $occupied.Server.ExclusiveAddressUse = $true
    $occupied.Start()
  }
  $blockedPort = [int]$occupied.LocalEndpoint.Port
  Assert-BrowserTest (-not (Test-BrowserPortAvailable $blockedPort)) '未识别真实占用端口'
  $launcher = Join-Path $testScripts 'open_sellersprite_login_browser.ps1'
  & $launcher -Port $blockedPort -ProfileDir $testProfile -BrowserPath $browserExecutable -SellerSpriteUrl 'about:blank' -AmazonUrl 'about:blank'
  if ($LASTEXITCODE -ne 0) { throw '浏览器启动器退出失败' }
  $config = Get-Content -LiteralPath $testConfig -Raw -Encoding UTF8 | ConvertFrom-Json
  Assert-BrowserTest ($config.port -ne $blockedPort) '没有自动更换占用端口'
  Assert-BrowserTest (Test-ProjectDevToolsEndpoint $config.port $testProfile) '备用端口没有连接到本测试配置'
  Assert-BrowserTest (-not (Test-ProjectDevToolsEndpoint $config.port (Join-Path $TestRoot 'other-profile'))) '错误复用了另一份浏览器配置'
  Assert-BrowserTest ($config.marker -eq '保留旧配置字段') '保存配置丢失原字段'
  Assert-BrowserTest ($config.cdp_url -eq "http://127.0.0.1:$($config.port)") '保存的连接地址与实际端口不一致'
  $launchCount = $global:WorkflowBrowserTestProcesses.Count
  & $launcher -Port $blockedPort -ProfileDir $testProfile -BrowserPath $browserExecutable -SellerSpriteUrl 'about:blank' -AmazonUrl 'about:blank'
  Assert-BrowserTest ($global:WorkflowBrowserTestProcesses.Count -eq $launchCount) '重复运行又打开了新窗口'
  $again = Get-Content -LiteralPath $testConfig -Raw -Encoding UTF8 | ConvertFrom-Json
  Assert-BrowserTest ($again.port -eq $config.port -and $again.profile_dir -eq $testProfile) '重复运行改变了端口或配置归属'
  Assert-BrowserTest (-not (Test-BrowserPortAvailable $blockedPort)) '错误结束了占用端口的其它服务'
  @{ status='passed'; occupied_port=$blockedPort; selected_port=$config.port; system_assigned_port=$systemPort; launches=$launchCount; profile_reused=$true } |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $TestRoot 'result.json') -Encoding UTF8
  Write-Host '[通过] 占用端口自动切换、浏览器归属校验、保存实际地址、原配置保留和重复启动复用。'
} catch {
  Write-Host ('测试启动参数：' + ($global:WorkflowBrowserTestArguments -join ' '))
  foreach ($item in @(Get-ProjectBrowserProcesses $testProfile)) {
    Write-Host ('测试进程：' + $item.ProcessId + ' ' + $item.CommandLine)
  }
  $errorLog = Join-Path $TestRoot 'browser.stderr.log'
  if (Test-Path -LiteralPath $errorLog) { Get-Content -LiteralPath $errorLog -Tail 12 | Write-Host }
  throw
} finally {
  # 只结束本测试实际启动的进程，不按名称批量关闭浏览器。
  foreach ($childId in $global:WorkflowBrowserTestProcesses) {
    Stop-Process -Id $childId -Force -ErrorAction SilentlyContinue
  }
  $occupied.Stop()
}
