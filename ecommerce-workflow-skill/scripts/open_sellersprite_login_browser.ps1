param(
  [int]$Port = 9223,
  [ValidateSet("US", "DE", "UK", "CA", "FR")]
  [string]$Site = "US",
  [string]$SellerSpriteUrl = "https://www.sellersprite.com/v3/keyword-miner/",
  [string]$AmazonUrl = "",
  [string]$BrowserPath = "",
  [string]$ProfileDir = ""
)

$ErrorActionPreference = "Stop"

$SkillDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$LocalBrowserConfig = Join-Path $SkillDir ".workflow-state\sellersprite\browser-config.json"

# 本机可以复用原项目配置；不把扩展和登录状态复制到新仓库。
if ([string]::IsNullOrWhiteSpace($ProfileDir) -and (Test-Path -LiteralPath $LocalBrowserConfig)) {
  $config = Get-Content -LiteralPath $LocalBrowserConfig -Raw -Encoding UTF8 | ConvertFrom-Json
  $ProfileDir = [string]$config.profile_dir
  if ([string]::IsNullOrWhiteSpace($ProfileDir)) {
    throw "本机浏览器配置缺少 profile_dir：$LocalBrowserConfig"
  }
}
if ([string]::IsNullOrWhiteSpace($ProfileDir)) {
  $ProfileDir = Join-Path $SkillDir ".workflow-state\sellersprite\.chrome-profile"
} elseif (-not (Test-Path -LiteralPath $ProfileDir -PathType Container)) {
  throw "指定的浏览器配置目录不存在，已停止以免误开空白配置：$ProfileDir"
}
$ProfileDir = [System.IO.Path]::GetFullPath($ProfileDir).TrimEnd('\')

function Resolve-BrowserPath {
  param([string]$ExplicitPath)

  # 与环境检查及图表渲染共用用户明确指定的浏览器路径。
  if ([string]::IsNullOrWhiteSpace($ExplicitPath) -and $env:CHROME_EXECUTABLE) {
    $ExplicitPath = $env:CHROME_EXECUTABLE
  }

  if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
    if (Test-Path -LiteralPath $ExplicitPath) {
      return (Resolve-Path -LiteralPath $ExplicitPath).Path
    }
    throw "指定的浏览器不存在：$ExplicitPath"
  }

  $candidates = @(
    (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
    (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe")
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return (Resolve-Path -LiteralPath $candidate).Path
    }
  }

  foreach ($commandName in @("chrome.exe", "msedge.exe")) {
    $command = Get-Command $commandName -ErrorAction SilentlyContinue
    if ($command) {
      return $command.Source
    }
  }

  throw "未找到 Google Chrome 或 Microsoft Edge。请先安装其中一个浏览器，或用 -BrowserPath 指定路径。"
}

function Test-DevToolsEndpoint {
  param([int]$DebugPort)
  try {
    $response = Invoke-RestMethod -Uri "http://127.0.0.1:$DebugPort/json/version" -TimeoutSec 2
    return -not [string]::IsNullOrWhiteSpace($response.webSocketDebuggerUrl)
  } catch {
    return $false
  }
}

if ([string]::IsNullOrWhiteSpace($AmazonUrl)) {
  $AmazonUrl = switch ($Site) {
    "DE" { "https://www.amazon.de/gp/css/homepage.html" }
    "UK" { "https://www.amazon.co.uk/gp/css/homepage.html" }
    "CA" { "https://www.amazon.ca/gp/css/homepage.html" }
    "FR" { "https://www.amazon.fr/gp/css/homepage.html" }
    default { "https://www.amazon.com/gp/css/homepage.html" }
  }
}

$Browser = Resolve-BrowserPath -ExplicitPath $BrowserPath
$listeners = @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)
if ($listeners.Count -gt 0 -and -not (Test-DevToolsEndpoint -DebugPort $Port)) {
  $owners = ($listeners | Select-Object -ExpandProperty OwningProcess -Unique) -join ", "
  throw "端口 $Port 已被其它程序占用（进程 ID：$owners），但不是可复用的项目浏览器。请关闭占用程序后重试。"
}
if ($listeners.Count -gt 0) {
  # 同一调试端口不代表同一份浏览器配置，防止连接到没有扩展的新窗口。
  foreach ($ownerId in @($listeners | Select-Object -ExpandProperty OwningProcess -Unique)) {
    $owner = Get-CimInstance Win32_Process -Filter "ProcessId = $ownerId"
    $profileMatch = [regex]::Match([string]$owner.CommandLine, '--user-data-dir(?:=|\s+)(?:"([^"]+)"|(\S+))')
    $actualProfile = if ($profileMatch.Groups[1].Success) { $profileMatch.Groups[1].Value } else { $profileMatch.Groups[2].Value }
    if (-not $profileMatch.Success -or [System.IO.Path]::GetFullPath($actualProfile).TrimEnd('\') -ine $ProfileDir) {
      throw "端口 $Port 的浏览器使用了另一份配置（进程 $ownerId）。请先关闭该项目窗口后重试。期望配置：$ProfileDir"
    }
  }
}

New-Item -ItemType Directory -Force -Path $ProfileDir | Out-Null

Start-Process -FilePath $Browser -ArgumentList @(
  "--remote-debugging-port=$Port",
  "--user-data-dir=`"$ProfileDir`"",
  "--new-window",
  $SellerSpriteUrl,
  $AmazonUrl
)

$deadline = (Get-Date).AddSeconds(20)
do {
  if (Test-DevToolsEndpoint -DebugPort $Port) {
    Write-Host "[完成] 项目专用浏览器已打开。"
    Write-Host "浏览器：$Browser"
    Write-Host "DevTools 端口：$Port"
    Write-Host "本机配置目录：$ProfileDir"
    Write-Host "使用已选定的项目配置；新电脑首次使用时安装卖家精灵插件，并登录卖家精灵和 Amazon。"
    exit 0
  }
  Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $deadline)

throw "浏览器已启动，但 20 秒内未检测到 DevTools 端口 $Port。请关闭该项目浏览器后重试。"

