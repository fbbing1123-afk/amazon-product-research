# 环境发现只读取本机路径，不安装依赖，不修改外部环境。
function Get-WorkflowRuntimeConfig {
  param([string]$SkillDir)
  $configPath = Join-Path $SkillDir ".workflow-state\runtime.json"
  if (Test-Path -LiteralPath $configPath) {
    return Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
  }
  return @{}
}

function Get-WorkflowPythonCandidates {
  param([string]$SkillDir, [string]$ExplicitPython = "")
  $config = Get-WorkflowRuntimeConfig $SkillDir
  $candidates = @($ExplicitPython, (Join-Path $SkillDir ".venv\Scripts\python.exe"), $config.python)
  foreach ($environmentDir in @($env:VIRTUAL_ENV, $env:CONDA_PREFIX)) {
    if ($environmentDir) {
      $candidates += Join-Path $environmentDir "Scripts\python.exe"
      $candidates += Join-Path $environmentDir "python.exe"
    }
  }
  $command = Get-Command python.exe -ErrorAction SilentlyContinue
  if ($command -and $command.Source -notmatch "WindowsApps") { $candidates += $command.Source }
  $launcher = Get-Command py.exe -ErrorAction SilentlyContinue
  if ($launcher) {
    $resolved = (& $launcher.Source -3 -c "import sys; print(sys.executable)" 2>$null | Out-String).Trim()
    if ($LASTEXITCODE -eq 0) { $candidates += $resolved }
  }
  $candidates += Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
  return @($candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -Unique)
}

function Test-WorkflowPython {
  param([string]$Python, [switch]$BaseOnly)
  if (-not $Python -or -not (Test-Path -LiteralPath $Python -PathType Leaf)) { return $false }
  $probe = "import sys; assert sys.version_info >= (3,10)"
  if (-not $BaseOnly) {
    $probe += "; import openpyxl, PIL, playwright, pandas, numpy; from importlib.metadata import version; v=lambda n:tuple(int(x) for x in version(n).split('.')[:2]); assert v('openpyxl') >= (3,1); assert v('playwright') >= (1,56); assert v('pandas') >= (2,2); assert v('numpy') >= (1,26)"
  }
  & $Python -c $probe *> $null
  return $LASTEXITCODE -eq 0
}

function Get-WorkflowNodeRoots {
  param([string]$SkillDir, [string]$ExplicitNodeModules = "")
  $config = Get-WorkflowRuntimeConfig $SkillDir
  return @(@($ExplicitNodeModules,
    (Join-Path $SkillDir "modules\product_research\scripts\chart-runtime\node_modules"),
    $config.node_modules, $env:CODEX_NODE_MODULES) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) } | Select-Object -Unique)
}

function Test-WorkflowNodeRoot {
  param([string]$ModulesPath)
  if (-not $ModulesPath -or -not (Get-Command node -ErrorAction SilentlyContinue)) { return $false }
  & node -e "const path=require('path'); for(const name of ['playwright','sharp','echarts']) require(path.join(process.argv[1],name));" $ModulesPath *> $null
  return $LASTEXITCODE -eq 0
}

function Find-WorkflowBrowser {
  $candidates = @($env:CHROME_EXECUTABLE)
  foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA)) {
    if ($base) {
      $candidates += Join-Path $base "Google\Chrome\Application\chrome.exe"
      $candidates += Join-Path $base "Microsoft\Edge\Application\msedge.exe"
    }
  }
  return $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
}
