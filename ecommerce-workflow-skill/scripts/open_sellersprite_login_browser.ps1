param(
  [int]$Port = 9223,
  [ValidateSet("US", "DE", "UK", "CA", "FR")]
  [string]$Site = "US",
  [string]$SellerSpriteUrl = "https://www.sellersprite.com/v3/keyword-miner/",
  [string]$AmazonUrl = "",
  [string]$BrowserPath = ""
)

$ErrorActionPreference = "Stop"

$SkillDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$ProfileDir = Join-Path $SkillDir ".workflow-state\sellersprite\.chrome-profile"

function Resolve-BrowserPath {
  param([string]$ExplicitPath)

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

New-Item -ItemType Directory -Force -Path $ProfileDir | Out-Null

Start-Process -FilePath $Browser -ArgumentList @(
  "--remote-debugging-port=$Port",
  "--user-data-dir=$ProfileDir",
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
    Write-Host "首次使用时，请在该浏览器中安装 softtime/sorftime 插件，并登录卖家精灵和 Amazon。"
    exit 0
  }
  Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $deadline)

throw "浏览器已启动，但 20 秒内未检测到 DevTools 端口 $Port。请关闭该项目浏览器后重试。"

