param(
  [ValidateRange(1, 65535)]
  [int]$Port = 9223,
  [ValidateSet("US", "DE", "UK", "CA", "FR")]
  [string]$Site = "US",
  [string]$SellerSpriteUrl = "https://www.sellersprite.com/v3/keyword-miner/",
  [string]$AmazonUrl = "",
  [string]$BrowserPath = "",
  [string]$ProfileDir = ""
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "browser_runtime.ps1")

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
$runningPort = Get-RunningProjectBrowserPort $ProfileDir
if ($runningPort) {
  # 已运行的同一配置优先复用，避免重复打开窗口或丢失插件登录。
  $Port = $runningPort
  Save-ProjectBrowserConfig $LocalBrowserConfig $ProfileDir $Browser $Port
  Write-Host "[复用] 项目浏览器已运行：http://127.0.0.1:$Port"
  Write-Host "本机配置目录：$ProfileDir"
  exit 0
}
if (@(Get-ProjectBrowserProcesses $ProfileDir).Count -gt 0) {
  throw "同一项目配置已有浏览器进程，但尚未提供可用的调试连接。请关闭使用这份配置的窗口后重试：$ProfileDir"
}
$preferredPort = $Port
$Port = Find-FreeBrowserPort $preferredPort
if ($Port -ne $preferredPort) {
  Write-Host "[自动切换] 端口 $preferredPort 被占用，使用空闲端口 $Port。"
}
New-Item -ItemType Directory -Force -Path $ProfileDir | Out-Null
Start-Process -FilePath $Browser -ArgumentList @(
  "--remote-debugging-address=127.0.0.1",
  "--remote-debugging-port=$Port",
  "--user-data-dir=`"$ProfileDir`"",
  "--new-window",
  $SellerSpriteUrl,
  $AmazonUrl
)

$deadline = (Get-Date).AddSeconds(20)
do {
  # 不能只检查端口是否响应；必须确认响应来自这份项目配置。
  if (Test-ProjectDevToolsEndpoint $Port $ProfileDir) {
    Save-ProjectBrowserConfig $LocalBrowserConfig $ProfileDir $Browser $Port
    Write-Host "[完成] 项目专用浏览器已打开。"
    Write-Host "浏览器：$Browser"
    Write-Host "CDP 地址：http://127.0.0.1:$Port"
    Write-Host "本机配置目录：$ProfileDir"
    Write-Host "端口已保存，后续脚本自动读取。新电脑首次使用时安装卖家精灵插件并登录。"
    exit 0
  }
  Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $deadline)
throw "浏览器启动后未在端口 $Port 提供本项目调试连接，配置记录未更新。请检查该项目窗口后重试。"
