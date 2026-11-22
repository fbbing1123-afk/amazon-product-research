param(
  [string]$Python = "",
  [string]$PackageManager = "",
  [switch]$RecreateVenv
)

$ErrorActionPreference = "Stop"

$SkillDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$VenvDir = Join-Path $SkillDir ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$ChartRuntimeDir = Join-Path $SkillDir "modules\product_research\scripts\chart-runtime"
$RequirementsFile = Join-Path $SkillDir "requirements.txt"
$EnvironmentCheck = Join-Path $SkillDir "scripts\check_environment.ps1"

function Invoke-CheckedCommand {
  param(
    [Parameter(Mandatory = $true)][string]$Executable,
    [Parameter(Mandatory = $true)][string[]]$Arguments,
    [Parameter(Mandatory = $true)][string]$Description
  )

  & $Executable @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "$Description 失败，退出代码：$LASTEXITCODE"
  }
}

function Resolve-PythonCommand {
  param([string]$ExplicitPython)

  $candidates = @()
  if (-not [string]::IsNullOrWhiteSpace($ExplicitPython)) {
    $candidates += [PSCustomObject]@{ Executable = $ExplicitPython; Prefix = @() }
  } else {
    $codexPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    $candidates += [PSCustomObject]@{ Executable = "py"; Prefix = @("-3") }
    $candidates += [PSCustomObject]@{ Executable = "python"; Prefix = @() }
    if (Test-Path -LiteralPath $codexPython) {
      $candidates += [PSCustomObject]@{ Executable = $codexPython; Prefix = @() }
    }
  }

  foreach ($candidate in $candidates) {
    try {
      $versionArgs = @($candidate.Prefix) + @("--version")
      $versionText = (& $candidate.Executable @versionArgs 2>&1 | Out-String).Trim()
      if ($LASTEXITCODE -eq 0 -and $versionText -match "Python\s+(\d+)\.(\d+)") {
        $major = [int]$Matches[1]
        $minor = [int]$Matches[2]
        if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 10)) {
          return $candidate
        }
      }
    } catch {
      continue
    }
  }

  throw "未找到 Python 3.10 或以上版本。请先安装 Python，并重新运行本脚本；也可用 -Python 指定 python.exe。"
}

function Resolve-PackageManager {
  param([string]$ExplicitPackageManager)

  if (-not [string]::IsNullOrWhiteSpace($ExplicitPackageManager)) {
    return $ExplicitPackageManager
  }
  foreach ($candidate in @("pnpm.cmd", "pnpm", "npm.cmd", "npm")) {
    if (Get-Command $candidate -ErrorAction SilentlyContinue) {
      return $candidate
    }
  }
  throw "未找到 pnpm 或 npm。请先安装 Node.js 18 或以上版本，再重新运行本脚本。"
}

Write-Host "[准备] 技能目录：$SkillDir"
Write-Host "[准备] Python 虚拟环境：$VenvDir"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "未找到 Node.js。请先安装 Node.js 18 或以上版本。"
}
$nodeVersionText = (& node --version 2>&1 | Out-String).Trim()
if ($nodeVersionText -notmatch "v(\d+)") {
  throw "无法识别 Node.js 版本：$nodeVersionText"
}
if ([int]$Matches[1] -lt 18) {
  throw "Node.js 版本过低：$nodeVersionText。需要 18 或以上版本。"
}

$pythonCommand = Resolve-PythonCommand -ExplicitPython $Python
$packageManager = Resolve-PackageManager -ExplicitPackageManager $PackageManager

if ($RecreateVenv -and (Test-Path -LiteralPath $VenvDir)) {
  Write-Host "[准备] 正在按要求重建 .venv..."
  Remove-Item -LiteralPath $VenvDir -Recurse -Force
}

if (-not (Test-Path -LiteralPath $VenvPython)) {
  Write-Host "[准备] 正在创建项目 Python 虚拟环境..."
  $venvArgs = @($pythonCommand.Prefix) + @("-m", "venv", $VenvDir)
  Invoke-CheckedCommand -Executable $pythonCommand.Executable -Arguments $venvArgs -Description "创建 Python 虚拟环境"
}

Write-Host "[准备] 正在安装 Python 依赖..."
Invoke-CheckedCommand -Executable $VenvPython -Arguments @("-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel") -Description "升级 pip"
Invoke-CheckedCommand -Executable $VenvPython -Arguments @("-m", "pip", "install", "-r", $RequirementsFile) -Description "安装 Python 依赖"

Write-Host "[准备] 正在安装 Playwright Chromium..."
$env:PLAYWRIGHT_BROWSERS_PATH = "0"
Invoke-CheckedCommand -Executable $VenvPython -Arguments @("-m", "playwright", "install", "chromium") -Description "安装 Playwright Chromium"

Write-Host "[准备] 正在安装图表运行时依赖..."
Push-Location $ChartRuntimeDir
try {
  if ($packageManager -match "pnpm") {
    Invoke-CheckedCommand -Executable $packageManager -Arguments @("install", "--frozen-lockfile") -Description "安装 Node.js 依赖"
  } else {
    Invoke-CheckedCommand -Executable $packageManager -Arguments @("install") -Description "安装 Node.js 依赖"
  }
} finally {
  Pop-Location
}

Write-Host "[准备] 正在执行安装后自检..."
& powershell -ExecutionPolicy Bypass -File $EnvironmentCheck -RequireInstalled
if ($LASTEXITCODE -ne 0) {
  throw "安装完成，但环境自检未通过。请把上面的失败项完整交给 Codex 排查。"
}

Write-Host ""
Write-Host "[完成] 本机运行环境已经准备好。"
Write-Host "[下一步] 运行以下命令打开项目专用浏览器："
Write-Host "powershell -ExecutionPolicy Bypass -File scripts\open_sellersprite_login_browser.ps1"
Write-Host "请只在该浏览器中安装 softtime/sorftime 插件，并登录卖家精灵和 Amazon。"

