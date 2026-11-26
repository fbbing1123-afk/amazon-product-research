# 项目浏览器的端口选择、归属校验和本机配置，不操作其它程序。
function Get-BrowserProfileFromCommandLine {
  param([string]$CommandLine)
  $match = [regex]::Match($CommandLine, '--user-data-dir(?:=|\s+)(?:"([^"]+)"|(\S+))')
  if (-not $match.Success) { return "" }
  $value = if ($match.Groups[1].Success) { $match.Groups[1].Value } else { $match.Groups[2].Value }
  return [System.IO.Path]::GetFullPath($value).TrimEnd('\')
}

function Get-ProjectBrowserProcesses {
  param([string]$ProfileDir)
  $expected = [System.IO.Path]::GetFullPath($ProfileDir).TrimEnd('\')
  return @(Get-CimInstance Win32_Process -Filter "Name = 'chrome.exe' OR Name = 'msedge.exe'" | Where-Object {
    (Get-BrowserProfileFromCommandLine ([string]$_.CommandLine)) -ieq $expected
  })
}

function Test-ProjectDevToolsEndpoint {
  param([int]$DebugPort, [string]$ProfileDir)
  $listeners = @(Get-NetTCPConnection -LocalPort $DebugPort -State Listen -ErrorAction SilentlyContinue)
  if (-not $listeners.Count) { return $false }
  $expected = [System.IO.Path]::GetFullPath($ProfileDir).TrimEnd('\')
  foreach ($ownerId in @($listeners | Select-Object -ExpandProperty OwningProcess -Unique)) {
    $owner = Get-CimInstance Win32_Process -Filter "ProcessId = $ownerId" -ErrorAction SilentlyContinue
    if (-not $owner -or (Get-BrowserProfileFromCommandLine ([string]$owner.CommandLine)) -ine $expected) {
      return $false
    }
  }
  try {
    $response = Invoke-RestMethod -Uri "http://127.0.0.1:$DebugPort/json/version" -TimeoutSec 2
    return -not [string]::IsNullOrWhiteSpace($response.webSocketDebuggerUrl)
  } catch { return $false }
}

function Get-RunningProjectBrowserPort {
  param([string]$ProfileDir)
  foreach ($process in @(Get-ProjectBrowserProcesses $ProfileDir)) {
    $match = [regex]::Match([string]$process.CommandLine, '--remote-debugging-port(?:=|\s+)(\d+)')
    if ($match.Success) {
      $candidate = [int]$match.Groups[1].Value
      if ($candidate -gt 0 -and (Test-ProjectDevToolsEndpoint $candidate $ProfileDir)) { return $candidate }
    }
  }
  return $null
}

function Test-BrowserPortAvailable {
  param([int]$DebugPort)
  $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $DebugPort)
  try {
    $listener.Server.ExclusiveAddressUse = $true
    $listener.Start()
    return $true
  } catch { return $false }
  finally { $listener.Stop() }
}

function Find-FreeBrowserPort {
  param([int]$PreferredPort = 9223)
  foreach ($candidate in @(@($PreferredPort) + @(9223..9243) | Select-Object -Unique)) {
    if (Test-BrowserPortAvailable $candidate) { return $candidate }
  }
  # 常用备用端口都被占用时，由系统分配一个空闲端口。
  $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
  try {
    $listener.Server.ExclusiveAddressUse = $true
    $listener.Start()
    return [int]$listener.LocalEndpoint.Port
  } finally { $listener.Stop() }
}

function Save-ProjectBrowserConfig {
  param([string]$ConfigPath, [string]$ProfileDir, [string]$BrowserPath, [int]$DebugPort)
  $settings = [ordered]@{}
  if (Test-Path -LiteralPath $ConfigPath) {
    $previous = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($property in $previous.PSObject.Properties) { $settings[$property.Name] = $property.Value }
  }
  $settings['profile_dir'] = $ProfileDir
  $settings['browser_path'] = $BrowserPath
  $settings['port'] = $DebugPort
  $settings['cdp_url'] = "http://127.0.0.1:$DebugPort"
  $settings['debugger_address'] = "127.0.0.1:$DebugPort"
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $ConfigPath) | Out-Null
  $temporary = "$ConfigPath.$([guid]::NewGuid().ToString('N')).tmp"
  try {
    $settings | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -LiteralPath $temporary -Destination $ConfigPath -Force
  } finally {
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary }
  }
}
