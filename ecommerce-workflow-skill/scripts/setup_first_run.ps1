param(
  [string]$Python = "",
  [string]$NodeModules = "",
  [string]$PackageManager = ""
)
$ErrorActionPreference = "Stop"
$SkillDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $PSScriptRoot "runtime_environment.ps1")
$StateDir = Join-Path $SkillDir ".workflow-state"
$ChartRuntimeDir = Join-Path $SkillDir "modules\product_research\scripts\chart-runtime"
function Invoke-Checked {
  param([string]$Executable, [string[]]$Arguments)
  & $Executable @Arguments
  if ($LASTEXITCODE -ne 0) { throw "命令失败，退出码$LASTEXITCODE；保留成功文件，修复后可重新运行。" }
}
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw "缺少Node.js，请先安装22或24；已有版本无需重装。" }
$nodeVersion = (& node --version | Out-String).Trim()
if ($nodeVersion -notmatch '^v(\d+)' -or [int]$Matches[1] -lt 18) { throw "Node.js版本不满足要求：$nodeVersion" }
if (-not (Find-WorkflowBrowser)) { throw "未找到Chrome或Edge；先安装浏览器或设置CHROME_EXECUTABLE。" }
if ($Python -and -not (Test-Path -LiteralPath $Python -PathType Leaf)) { throw "指定Python不存在：$Python" }
if ($NodeModules -and -not (Test-Path -LiteralPath $NodeModules -PathType Container)) { throw "指定Node依赖目录不存在：$NodeModules" }
$pythonCandidates = @(Get-WorkflowPythonCandidates $SkillDir $Python)
$readyPython = $pythonCandidates | Where-Object { Test-WorkflowPython $_ } | Select-Object -First 1
if ($readyPython) {
  Write-Host "[复用] Python依赖已可用，不执行pip安装：$readyPython"
} else {
  $basePython = $pythonCandidates | Where-Object { Test-WorkflowPython $_ -BaseOnly } | Select-Object -First 1
  if (-not $basePython) { throw "未找到Python 3.10以上版本，推荐3.12；可用-Python指定已有解释器。" }
  $VenvDir = Join-Path $SkillDir ".venv"
  $readyPython = Join-Path $VenvDir "Scripts\python.exe"
  if (-not (Test-Path -LiteralPath $readyPython)) {
    Write-Host "[安装] 创建项目内Python环境：$VenvDir"
    Invoke-Checked $basePython @('-m', 'venv', $VenvDir)
  }
  # 只向项目虚拟环境安装，外部解释器及系统全局依赖保持不变。
  if (-not (Test-WorkflowPython $readyPython -BaseOnly)) { throw "项目.venv不可用，请先检查原环境，不自动删除或覆盖。" }
  $pipCache = Join-Path $StateDir "pip-cache"
  Invoke-Checked $readyPython @('-m', 'pip', 'install', '--cache-dir', $pipCache, '-r', (Join-Path $SkillDir 'requirements.txt'))
}
$readyNode = Get-WorkflowNodeRoots $SkillDir $NodeModules | Where-Object { Test-WorkflowNodeRoot $_ } | Select-Object -First 1
if ($readyNode) {
  Write-Host "[复用] 图表依赖已可用，不执行Node依赖安装：$readyNode"
} else {
  $manager = $PackageManager
  if (-not $manager) {
    foreach ($candidate in @('npm.cmd', 'pnpm.cmd')) {
      if (Get-Command $candidate -ErrorAction SilentlyContinue) { $manager = $candidate; break }
    }
  }
  if (-not $manager) { throw "未找到npm或pnpm，请检查Node.js安装。" }
  Write-Host "[安装] 图表依赖写入项目chart-runtime目录"
  Push-Location $ChartRuntimeDir
  try {
    if ($manager -match 'pnpm') {
      Invoke-Checked $manager @('install', '--frozen-lockfile', '--store-dir', (Join-Path $StateDir 'pnpm-store'))
    } else {
      Invoke-Checked $manager @('install', '--cache', (Join-Path $StateDir 'npm-cache'), '--no-audit', '--no-fund')
    }
  } finally { Pop-Location }
  $readyNode = Join-Path $ChartRuntimeDir 'node_modules'
}
# 新版网页使用现有项目Chrome，离线图表也复用已安装浏览器，无需另下载Chromium。
if (-not (Test-WorkflowPython $readyPython) -or -not (Test-WorkflowNodeRoot $readyNode)) { throw "准备后依赖仍不可用，未写入成功配置。" }
New-Item -ItemType Directory -Force -Path $StateDir | Out-Null
$config = [ordered]@{ python = [System.IO.Path]::GetFullPath($readyPython); node_modules = [System.IO.Path]::GetFullPath($readyNode) }
$config | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $StateDir 'runtime.json') -Encoding UTF8
Invoke-Checked 'powershell.exe' @('-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'check_environment.ps1'), '-RequireInstalled')
Write-Host "[完成] 环境可用，记录于.workflow-state/runtime.json。后续使用其中python路径，并将CODEX_NODE_MODULES设为记录的node_modules。"
Write-Host "[下一步] 打开项目专用浏览器，安装卖家精灵插件，分别登录插件、卖家精灵网页和Amazon。"
