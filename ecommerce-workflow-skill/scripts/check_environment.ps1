param(
  [switch]$RequireInstalled,
  [string]$Python = "",
  [string]$NodeModules = ""
)
$ErrorActionPreference = "Stop"
$SkillDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $PSScriptRoot "runtime_environment.ps1")
$failures = New-Object System.Collections.Generic.List[string]
function Write-Check {
  param([string]$Name, [bool]$Passed, [string]$Detail)
  Write-Host ("[{0}] {1}：{2}" -f $(if ($Passed) { "通过" } else { "失败" }), $Name, $Detail)
  if (-not $Passed) { $script:failures.Add($Name) }
}
Write-Host "亚马逊选品自动化｜当前测试版环境检查"
Write-Host "技能目录：$SkillDir"
$requiredFiles = @(
  "SKILL.md", "agents\openai.yaml", "requirements.txt", "docs\delivery_checks.md",
  "scripts\build_amazon_collection_link.py", "scripts\run_review_prefetch.py",
  "scripts\open_sellersprite_login_browser.ps1", "scripts\browser_runtime.ps1",
  "modules\shared\browser_runtime.py", "modules\shared\browser_runtime.mjs",
  "modules\shared\amazon_access.py",
  "modules\competitor_analysis\fill_review_batch.py", "resources\asin_collection_sites.json",
  "modules\product_research\scripts\chart-runtime\package.json",
  "modules\product_research\scripts\chart-runtime\pnpm-lock.yaml",
  "modules\product_research\scripts\chart-runtime\vendor\echarts.min.js"
)
$missing = @($requiredFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $SkillDir $_)) })
Write-Check "技能文件" ($missing.Count -eq 0) $(if ($missing.Count) { $missing -join ', ' } else { "齐全" })
$nodeCommand = Get-Command node -ErrorAction SilentlyContinue
$nodeVersion = if ($nodeCommand) { (& node --version | Out-String).Trim() } else { "未找到" }
$nodeReady = $nodeVersion -match '^v(\d+)' -and [int]$Matches[1] -ge 18
Write-Check "Node.js" $nodeReady "$nodeVersion，推荐22或24"
$browser = Find-WorkflowBrowser
Write-Check "浏览器" ([bool]$browser) $(if ($browser) { $browser } else { "请安装Chrome，或指定CHROME_EXECUTABLE" })
$readyPython = Get-WorkflowPythonCandidates $SkillDir $Python | Where-Object { Test-WorkflowPython $_ } | Select-Object -First 1
if ($readyPython) { Write-Check "可复用Python依赖" $true $readyPython }
elseif ($RequireInstalled) { Write-Check "Python依赖" $false "未找到可用环境，可用-Python指定已有解释器" }
else { Write-Host "[待准备] Python依赖：先指定或确认已有环境，确实缺少再运行setup_first_run.ps1" }
$readyNode = Get-WorkflowNodeRoots $SkillDir $NodeModules | Where-Object { Test-WorkflowNodeRoot $_ } | Select-Object -First 1
if ($readyNode) { Write-Check "可复用图表依赖" $true $readyNode }
elseif ($RequireInstalled) { Write-Check "图表依赖" $false "未找到可导入的playwright、sharp、echarts" }
else { Write-Host "[待准备] 图表依赖：可用-NodeModules指定已有目录，确实缺少才安装到项目内" }
if ($failures.Count) { Write-Host "未通过：$($failures -join '、')"; exit 1 }
Write-Host "检查结束；仍须核验插件和网页登录、站点邮编及输入模板。此检查不访问业务网站。"
exit 0
