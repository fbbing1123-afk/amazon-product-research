param(
  [switch]$RequireInstalled
)

$ErrorActionPreference = "Continue"
$SkillDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$VenvPython = Join-Path $SkillDir ".venv\Scripts\python.exe"
$ChartRuntimeDir = Join-Path $SkillDir "modules\product_research\scripts\chart-runtime"
$failures = New-Object System.Collections.Generic.List[string]

function Add-CheckResult {
  param(
    [string]$Name,
    [bool]$Passed,
    [string]$Detail
  )
  $status = if ($Passed) { "通过" } else { "失败" }
  Write-Host ("[{0}] {1}：{2}" -f $status, $Name, $Detail)
  if (-not $Passed) {
    $script:failures.Add($Name)
  }
}

function Find-Browser {
  $candidates = @(
    (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
    (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe"),
    (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe")
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  return $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}

Write-Host "亚马逊选品自动化 Skill 环境自检"
Write-Host "技能目录：$SkillDir"
Write-Host ""

$requiredFiles = @(
  "SKILL.md",
  "agents\openai.yaml",
  "requirements.txt",
  "scripts\run_workflow.py",
  "modules\product_research\scripts\chart-runtime\package.json",
  "modules\product_research\scripts\chart-runtime\pnpm-lock.yaml",
  "modules\product_research\scripts\chart-runtime\vendor\echarts.min.js"
)
$missingFiles = @($requiredFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $SkillDir $_)) })
Add-CheckResult -Name "技能文件" -Passed ($missingFiles.Count -eq 0) -Detail $(if ($missingFiles.Count -eq 0) { "关键文件齐全" } else { "缺少：$($missingFiles -join ', ')" })

$nodeCommand = Get-Command node -ErrorAction SilentlyContinue
$nodeVersion = if ($nodeCommand) { (& node --version 2>&1 | Out-String).Trim() } else { "未找到" }
$nodeReady = $nodeVersion -match "v(\d+)" -and [int]$Matches[1] -ge 18
Add-CheckResult -Name "Node.js" -Passed $nodeReady -Detail "$nodeVersion（需要 18 或以上）"

$browser = Find-Browser
Add-CheckResult -Name "浏览器" -Passed ($null -ne $browser) -Detail $(if ($browser) { $browser } else { "未找到 Chrome 或 Edge" })

$venvExists = Test-Path -LiteralPath $VenvPython
if ($venvExists) {
  $pythonVersion = (& $VenvPython --version 2>&1 | Out-String).Trim()
  & $VenvPython -c "import openpyxl, PIL, playwright; print('依赖可导入')" *> $null
  $pythonReady = $LASTEXITCODE -eq 0
  Add-CheckResult -Name "Python 虚拟环境" -Passed $pythonReady -Detail "$pythonVersion；依赖检查 $(if ($pythonReady) { '通过' } else { '失败' })"
} elseif ($RequireInstalled) {
  Add-CheckResult -Name "Python 虚拟环境" -Passed $false -Detail "未找到 .venv\Scripts\python.exe"
} else {
  Write-Host "[待安装] Python 虚拟环境：运行 scripts\setup_first_run.ps1 后生成"
}

$nodeDependencies = @("playwright", "sharp", "echarts")
$missingNodeDependencies = @($nodeDependencies | Where-Object { -not (Test-Path -LiteralPath (Join-Path $ChartRuntimeDir "node_modules\$_")) })
if ($missingNodeDependencies.Count -eq 0) {
  Add-CheckResult -Name "图表运行时" -Passed $true -Detail "Node.js 依赖齐全"
} elseif ($RequireInstalled) {
  Add-CheckResult -Name "图表运行时" -Passed $false -Detail "缺少：$($missingNodeDependencies -join ', ')"
} else {
  Write-Host "[待安装] 图表运行时：运行 scripts\setup_first_run.ps1 后生成"
}

Write-Host ""
if ($failures.Count -gt 0) {
  Write-Host "自检未通过：$($failures -join '、')"
  exit 1
}

if ($RequireInstalled) {
  Write-Host "自检通过，可以运行工作流。"
} else {
  Write-Host "基础条件通过。若显示待安装，请继续运行 scripts\setup_first_run.ps1。"
}
exit 0

