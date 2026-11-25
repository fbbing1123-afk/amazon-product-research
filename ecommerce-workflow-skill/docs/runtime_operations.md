# 运行、环境与检查点

命令均从Skill根目录运行。先读取 `.workflow-state/runtime.json` 的 `python`、`node_modules`，复用其中路径；没有记录时检查当前项目、激活环境及用户指定环境。不要沿用作者路径。PowerShell脚本是UTF-8 BOM编码，兼容Windows PowerShell读取中文。

## 初始化

```powershell
powershell -ExecutionPolicy Bypass -File scripts\check_environment.ps1
powershell -ExecutionPolicy Bypass -File scripts\setup_first_run.ps1
powershell -ExecutionPolicy Bypass -File scripts\check_environment.ps1 -RequireInstalled
powershell -ExecutionPolicy Bypass -File scripts\open_sellersprite_login_browser.ps1 -Site US
```

已有Python但不在PATH时向检查/准备脚本传 `-Python "本机python.exe绝对路径"`；图表依赖可传 `-NodeModules "已有node_modules目录"`。脚本先检查依赖可导入性，可用即复用，不执行安装；缺项才在本项目 `.venv`、`chart-runtime/node_modules` 及 `.workflow-state` 下载缓存内补齐，不向外部解释器或全局Node安装。已有同机项目浏览器配置时向启动器传 `-ProfileDir "本机已有目录"`，不要拷贝其它人的配置。浏览器启动器提供US/DE/UK/CA/FR；本地集合链接另外支持12站，两者不等同。

安装脚本只安装项目依赖，不自动安装系统Python、Node、Chrome或替用户登录。不会自动删除或重建已有环境。若缺软件，由AI明确指出缺项并完成适当的本机安装后继续。浏览器优先复用Chrome；自定义安装位置可设 `CHROME_EXECUTABLE`，离线图表同样复用它，不额外下载Chromium。

## 关键词阶段

```powershell
$runtime = Get-Content -LiteralPath '.workflow-state/runtime.json' -Raw -Encoding UTF8 | ConvertFrom-Json
$env:CODEX_NODE_MODULES=$runtime.node_modules
$env:PYTHONPATH='.;modules'
$env:PYTHONUTF8='1'
& $runtime.python scripts\run_workflow.py `
  --source '用户本批Excel绝对路径' `
  --site US --amazon-postal-code 10001 `
  --steps keyword_research `
  --keyword-ppc-bid-mode latest-week `
  --keyword-backfill-debugger-address '127.0.0.1:9223' `
  --keyword-trends-debug-port 9223 --keyword-trends-use-existing-browser
```

其余当前阶段使用 `new_workflow_trial.md` 的独立入口。不要去掉 `--steps keyword_research` 或往该命令加入旧市场、类目、评论步骤。复用手动关键词导出时先核对文件、关键词和站点，再使用该脚本对应导入参数。

## 监控与恢复

- 启动前记录批次、范围、数量、源文件指纹和已有检查点；各阶段记录开始/结束、耗时、结果与失败项。没有实测依据不编造预计耗时。
- 约30秒报告真实进展；下载状态持续变化可继续等模块超时，单纯心跳不等于进度。连续60秒无日志、文件或页面变化，检查具体卡点；再次60秒无进展则停止当前子任务，保留成果。
- CDP单次连接20秒，先查本机 `/json/version`，只允许恢复同一项目配置重连一次，不结束其它Chrome或Python进程。
- 普通失败最多定点补跑一次；类目按其独立规则不自动反复刷新。验证、登录或账号权限失败交用户完成，不把失败转换为成功的空数据。
- 下载100%不等于文件保存；只在文件完成并校验归属后关闭商品页。失败页保留，续跑只传失败ASIN和必要参数。
- 基于检查点恢复；新源文件指纹或商品集合变化后重新准备对应分析，不能套旧JSON。不重复截图、下载或修改已成功部分。
- `connect_over_cdp`连接结束仅断开客户端/退出Playwright上下文，不调用 `browser.close()`。本机自己 `launch` 的离线图表浏览器可正常关闭。

关键词总入口监控文件可用 `scripts/inspect_workflow_monitor.py` 和 `scripts/monitor_workflow.py` 检查；预取有自己的 `manifest.json`、`events.jsonl`。主智能体串联的其它阶段也在本批目录记录结果、异常和输出路径。
