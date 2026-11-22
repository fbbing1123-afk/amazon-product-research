---
name: ecommerce-workflow-skill
description: 亚马逊选品自动化本地工作流 Skill。用于运行、维护或分发一个“单个 Excel 工作簿贯穿全流程”的模块化流程，按需串联关键词调研、类目调研、市场分析、商品评论抓取和子智能体竞品分析，并输出产品调研 Excel。
---

# 亚马逊选品自动化 SKILL

## 这个 Skill 是做什么的

这个项目用于把一个亚马逊/卖家精灵数据源 Excel 工作簿，按固定流程加工成产品调研和市场分析报告。

面向同事的安装、登录、运行、监控和打包步骤见：[docs/使用指南.docx](docs/使用指南.docx)；同步文本版见 [docs/使用指南.md](docs/使用指南.md)。本文是智能体执行规则；分发项目时两份文档都要保留。

## 分发版首次启动协议

当技能目录位于新电脑，或用户说“第一次使用”“安装技能”“帮我配置环境”时，主智能体必须按以下顺序执行：

1. 先读取 [docs/使用指南.md](docs/使用指南.md)，确认技能安装路径和基础条件。
2. 运行 `powershell -ExecutionPolicy Bypass -File scripts\check_environment.ps1`。不要把系统的 WindowsApps `python.exe` 占位程序当成有效 Python。
3. `.venv\Scripts\python.exe` 或图表目录 `node_modules` 不存在时，运行 `powershell -ExecutionPolicy Bypass -File scripts\setup_first_run.ps1`；必须监管到安装后自检通过，不能只启动安装命令就结束。
4. 运行 `scripts\open_sellersprite_login_browser.ps1 -Site <站点>` 打开项目专用浏览器。浏览器脚本会自动选择 Chrome 或 Edge；不得关闭或覆盖其它程序占用的端口。
5. 需要用户安装插件、登录或处理验证码时，明确暂停在该步骤；不得索要或记录用户名、密码、Cookie、验证码。
6. 用户确认登录和插件可用后，再按请求检查 Excel 并运行工作流。首次安装请求本身不等于授权运行任意 Excel。

分发给同事时，优先让同事解压后直接在 Codex 中打开 `ecommerce-workflow-skill` 文件夹，并在这个文件夹中开启对话。用户只要要求 AI 先阅读当前目录的 `SKILL.md` 和 `docs/使用指南.md`，主智能体就应按本协议接管首次准备和后续工作流；不要要求没有编程经验的用户先手动复制到隐藏目录或输入 PowerShell 命令。需要长期安装为全局 Skill 时，才可复制到 `%USERPROFILE%\.codex\skills\ecommerce-workflow-skill` 并通过 `$ecommerce-workflow-skill` 调用。

默认完整流程：

```text
keyword_research -> category_research -> market_analysis -> review_research -> competitor_analysis
```

其中前四段由 `scripts\run_workflow.py` 串联，`competitor_analysis` 是评论调研之后的 Agent 语义分析阶段。主智能体完成 handoff JSON 后，必须立即调用 `modules\competitor_analysis\run.py`，在同一监控上下文中连续完成 Excel 回填、差评透视和切片器，不能让完整流程停在交接点等待下一轮对话。默认完整流程的评论对象只取“竞品分析”页列出的竞品 ASIN，不得读取并抓取整个数据源的全部 ASIN。用户明确要求只运行前三段或单独模块时，才通过 `--steps` 指定部分流程。未明确传入 `--site` 时，脚本会从数据源 Amazon 域名自动识别站点；数据源混有多个站点时必须明确传入 `--site`。

核心原则：同一个 Excel 工作簿在各步骤之间继续流转。后续步骤基于前一步输出的工作簿继续处理，而不是各自生成互不相关的文件。每个业务功能都必须独立放在 `modules\功能名\` 目录下，方便后续维护和新增模块。

## 模块独立运行边界

- `keyword_research` 是对外的关键词调研模块，内部依次执行关键词导出和关键词趋势。旧入口 `keyword_backfill`、`keyword_trends` 仍可分别单独运行，也可以按顺序组合运行；不能在已经传入 `keyword_research` 时再次混入这两个旧入口。
- `市场调研` 必须按完整三小节交付：第 1 小节“关键词搜索量及趋势”、第 2 小节“所属平台分类及常用分类节点”、第 3 小节“数据分析概览”。用户要求完成市场调研时，关键词、趋势截图、趋势旺季、站内结果、淡季/月搜索量、旺季/月搜索量、旺季竞价和关键词超链接都属于必填项，不能只完成类目截图或第 3 小节后就报告完成。
- `category_research` 是类目调研模块，可以单独运行；它需要当前工作簿有数据源和市场调研模板。`product_research` 只是兼容旧名称，与 `category_research` 二选一。
- `market_analysis` 是市场分析模块，可以单独运行；它只依赖数据源页，默认同时生成市场调研第 3 小节的 9 张图表和市场分析页。
- 单独运行“市场调研”工作表时，必须把第 1、2、3 小节作为一个完整交付单元处理，不能只运行关键词、类目截图或第 3 小节中的某一部分。关键词趋势图必须逐行锚定到“搜索趋势”列；类目表现截图必须按最小类目顺序锚定到当前类目数据行的“类目表现”列；第 3 小节的 9 张图表必须按当前标题行重新锚定到对应位置。关键词或类目扩行后，必须重新读取表格标题和数据行再校正浮动图片锚点，不能沿用扩行前的旧行号。完成后逐张核对图片内容、顺序、目标单元格和有效像素，并检查没有重复或错位图片后，才能报告市场调研完成。
- `review_research` 是评论调研模块，可以单独运行；它需要数据源中有 ASIN，或通过 `--review-asin` 显式传入 ASIN。工作簿存在“竞品分析”页时，脚本默认优先读取该页的竞品 ASIN，避免把整个数据源当成评论对象；只有没有竞品分析页时才回退读取数据源 ASIN。它不依赖卖家精灵，但依赖 Amazon 登录态和目标站点邮编。
- `competitor_analysis` 是独立的 Agent 语义模块，不作为 `scripts/run_workflow.py --steps` 的普通浏览器步骤。主智能体负责生成语义 handoff，随后立即调用 `modules\competitor_analysis\run.py`，由一条受监控命令连续完成竞品回填、差评透视和切片器。
- 多个脚本模块可以按 `keyword -> category -> market -> review` 顺序组合运行；只传其中一个或其中几个时，流程只执行所传模块，结果写入桌面对应的部分流程文件。启动前会做只读前置校验，输入页签或表头不足时立即报错，不会打开浏览器后才发现无法处理。

## 快速执行协议

每次运行前先判断请求类型，禁止不加判断地从头执行：

1. **新案例完整流程**：输入是新模板或新数据源，并且用户明确说“完整流程”，运行 `keyword_research -> category_research -> market_analysis -> review_research`，随后立即委派专用子智能体完成 `competitor_analysis` 和差评透视；不得在前三段完成后报告完整流程结束。
2. **从现有工作簿继续**：用户说“数据源已填好”“接着运行”时，先检查现有 sheet 和中间结果，从第一个未完成步骤开始。已完成的关键词导出、趋势、类目截图或图表不得重跑。
3. **补齐或重试**：用户说“补上”“补齐”“重试”时，先列出缺失的关键词、类目、ASIN、截图或字段，只处理缺失项。除非现有结果整体不可用，否则不得重跑完整列表。
4. **局部格式修复**：字体、颜色、边框、列宽、图片尺寸、单个图表或少量单元格问题，只修改对应工作簿区域；不得重新抓网页、重新分析评论或重建整份报告。
5. **用户指定工作表即按完整交付处理**：用户要求“完成某个 sheet”或某个功能模块时，必须完成该 sheet 的全部应填字段、中文翻译、图片、超链接、图表、边框和既有格式。缺失翻译或其它后处理属于本次任务的一部分，主智能体应自动读取 missing JSON、补齐缓存并回写工作簿，不得要求用户另行提醒“补翻译”或“补格式”。

完整流程中途因登录、验证码或浏览器短暂失联而只续跑 `review_research` 时，主智能体必须使用 `scripts\\run_workflow.py --continue-full-workflow --steps review_research`。这个标记只用于完整流程续跑：脚本会写入 `workflow_handoff_required`，主智能体随后必须立即调用竞品分析后处理；不带标记的 `review_research` 仍是独立模块，完成后可以正常写入 `workflow_finished`，不能自动触发竞品分析。

完整流程的唯一结束判定：

1. `run_workflow.py` 完成四个脚本阶段后，只得到 `agent_handoff_workbook`，不是最终完成状态；主智能体必须在同一轮立即准备语义 handoff，并调用 `modules\competitor_analysis\run.py` 继续处理。
2. 先处理类目缺失翻译 JSON；只回填缺失类目，不重跑类目截图。
3. 再处理 Critical reviews 缺失翻译和缺失痛点分组 JSON；只回填缺失内容，不重抓评论。
4. 随后完成竞品商品卡、基础字段、对比维度、评价分析、品类总结和改进方式。
5. 最后生成差评透视及 3 个橙色切片器。切片器完成后禁止再用 `openpyxl` 保存。
6. 只有最终结构、图片、链接、边框、评论数量、翻译、痛点、公式和视觉验收全部通过，才可向用户报告“完整流程完成”。

### 模块硬门禁

所有模块输出都必须经过 `modules\shared\quality_gates.py`。`scripts\run_workflow.py` 会在关键词、类目、市场分析和评论阶段后自动执行对应门禁；任一失败立即终止，不写完成状态。

```powershell
python -m shared.quality_gates --step keyword_research --workbook "关键词检查点.xlsx"
python -m shared.quality_gates --step category_research --workbook "类目检查点.xlsx"
python -m shared.quality_gates --step market_analysis --workbook "市场分析检查点.xlsx"
python -m shared.quality_gates --step review_research --workbook "评论检查点.xlsx"
python -m shared.quality_gates --step competitor_basics --workbook "竞品基础.xlsx"
python -m shared.quality_gates --step competitor_review_analysis --workbook "竞品评价完成.xlsx"
```

最终 WPS 交付必须使用带门禁的入口，不能直接调用普通转换脚本：

```powershell
python scripts\finalize_delivery.py `
  --original "原始模板.xlsx" `
  --floating-workbook "最终浮动版.xlsx" `
  --output "最终交付.xlsx" `
  --embed-extra-sheet "产品名-市场分析" `
  --embed-extra-sheet "产品名-竞品分析"
```

如本次跳过竞品评价分析，不传 `--allow-evaluation`；执行评价分析后才可传该参数。门禁会拒绝以下结果：类目链接不是 BSR、截图被实际缩小导致模糊、图片空白、红色规格丢失、评价区错误写入、公式错误或 ZIP 损坏。中间文件检查通过不能替代最终文件检查。

执行时遵守以下时间和复用规则：

- 先读取工作簿结构、目标站点、已有缓存和最近一个可用中间工作簿，再启动浏览器任务。
- 启动前先报告本次实际要处理的关键词数、去重类目数、目标 ASIN 数和待补缺失项数；数量异常时先复核，不能直接进入长时间批处理。
- 项目专用浏览器和登录态有效时直接复用，同一阶段只做一次前置登录/邮编检查；只有 CDP 失效时才重启浏览器。
- `category_research` 先按节点 ID 去重，类目表中的每一个类目都必须截图；每个类目等待目标插件区域出现最多 30 秒，区域出现后固定等待满 20 秒再截图。
- 翻译和痛点归一化放在确定性脚本之后，只处理缺失 JSON 中的新内容；缓存已命中的内容不得重新翻译。
- 每个成功步骤的输出都是续跑检查点。后续失败时从该检查点继续，不回到原始模板重做已成功步骤。
- 运行中的命令必须持续监管到结束或明确阻塞；每约 30 秒报告当前步骤、已完成数量和剩余数量，不用长时间静默等待。
- 浏览器或子进程连续 60 秒没有新增日志、页面变化或完成数量时，立即检查终端与目标页面并处理具体卡点；不得在未知状态下盲等 10 分钟。
- 当前用户已声明本机环境可用；日常运行不得为了执行案例新建 `.venv`。直接使用现有 Python，或通过 `--python` 指定可用解释器。后文“新电脑初始化”只适用于首次分发到尚未配置依赖的新电脑。
- 分发给没有编程能力的同事时，首次执行本 Skill 先检查 `.venv/Scripts/python.exe` 和 `modules/product_research/scripts/chart-runtime/node_modules`；任一不存在就由主智能体执行 `scripts/setup_first_run.ps1`，完成后再打开项目专用浏览器。不得改为在全局 Python 或全局 Node 中安装项目依赖。

### 防止完整流程长时间失控（强制）

完整流程不能只启动一个长命令后被动等待。主智能体必须把关键词、类目、市场分析、评论、竞品分析、差评透视和最终验收作为可观察的阶段管理：

1. **运行前登记工作量**：记录开始时间、输入文件、站点、关键词数、去重后的类目数、竞品 ASIN 数、Positive/Critical 任务数和已有检查点，并根据实际数量与最近一次同类运行记录给出预计耗时区间。没有可参考记录时应明确写“暂无可靠预计”，不能编造精确时间，也不能在未确认数量时直接开始批处理。
2. **每阶段单独计时**：至少记录阶段名称、开始时间、结束时间、耗时、成功数量、失败数量和输出检查点。每约 30 秒只报告真实进展；“仍在等待”不算进展。最终交付时报告各阶段耗时和总耗时，不能只凭印象估算。
3. **CDP 连接硬限制**：连接 `127.0.0.1:9223` 单次最多等待 20 秒。失败后先检查 `/json/version`，只允许重启同一项目 Chrome 配置并重连 1 次；再次失败就终止当前阶段并明确报告阻塞，禁止连续使用 180 秒连接等待。
4. **CDP 连接关闭规则**：通过 `connect_over_cdp` 接入项目专用浏览器后，只能调用 `disconnect()` 断开自动化客户端；禁止调用 `browser.close()` 关闭远程浏览器。只有脚本自己通过 `launch()` 启动的本地浏览器才允许调用 `close()`。
4. **未知卡顿硬停止**：同一浏览器或子进程连续 60 秒没有新增日志、页面变化、文件变化或完成数量时立即诊断；处理后再连续 60 秒仍无进展，只终止本次启动的当前进程，从最近检查点续跑。不得结束其它无关 Python/Chrome 进程，也不得回到原始模板重跑成功阶段。卖家精灵后台导出等“页面能读取到明确处理中状态、脚本持续输出轮询日志、模块自身已有确定超时”的任务不属于未知卡顿，可以等到该模块规定的超时；状态消失、页面失联或超过模块超时后必须停止。
5. **失败项最多补跑 1 次**：除非具体模块另有更严格规则，每个任务项（关键词、商品卡或 `(ASIN, 评论类型)` 等）正常执行 1 次，失败且结果仍缺失时只允许定点补跑 1 次；仍失败就记录原因并继续其它项目。`category_research` 保持现有更严格规则，类目最多 5 个、每个目标区域最多等待 30 秒且不自动重试。
6. **评论任务封顶**：每个 ASIN 分别建立一个 `Positive reviews` 任务和一个 `Critical reviews` 任务，正常情况下各执行 1 次；只有对应的 `(ASIN, 评论类型)` 任务失败且结果仍缺失时，才按上一条定点补跑。Positive 候选最多 50 条、Critical 候选最多 100 条，每组最多 10 个加载批次。达到任一上限、按钮消失或状态明确后立即停止，不能为了凑满过滤后的可信评论继续加载，也不能重复抓取已经数量对平且状态合格的任务。
7. **局部失败只做局部恢复**：趋势缺图只补对应关键词，类目缺图只补对应节点，评论失败只补对应 `(ASIN, 评论类型)`，翻译或归一化缺失只处理 missing JSON，格式问题只修对应单元格或图片。禁止因最终验收发现一个局部问题而重跑完整流程。
8. **阶段预算异常即停下来判断**：有可靠预计区间时，实际阶段耗时达到预计上限的 2 倍必须先检查日志、页面、输出文件和进程状态，再决定定点续跑或报告阻塞；没有可靠预计时按第 4 条的无进展规则控制。不能让任务在没有新产出的情况下持续数小时。
9. **漏填验收是交付前置条件**：脚本退出不等于步骤完成。必须按本次请求的目标工作表逐行检查所有应填的可见字段、超链接、图片对象及图片有效像素、图表对象或冻结图表；合并单元格和模板预留的未使用层级可排除，但其它非预期空值必须先定点补齐或明确报告，不能直接交付。
10. **市场调研三小节验收**：对 `市场调研` 的请求必须分别检查第 1、2、3 小节；第 1 小节至少核对关键词单元格、Amazon 超链接、每个关键词的趋势图片、趋势旺季、站内结果、淡季搜索量、旺季搜索量和旺季竞价；第 2 小节核对最小类目名称、超链接和中文翻译；第 3 小节核对 9 张图表。任一小节缺项时，只补缺失项，不得把“市场分析”已完成当成“市场调研”已完成。
11. **运行监控必须落地**：每次工作流运行都在缓存目录生成 `{name}_workflow_monitor.log` 和 `{name}_workflow_monitor.jsonl`。`.log` 供人工快速查看，`.jsonl` 记录工作流命令、子进程输出、文件变化、评论 ASIN/筛选/批次、页面导航、点击 `Show 10 more`、异常和耗时。连续 60 秒没有日志或被监控文件变化时自动写入 `inactivity_warning`；连续 180 秒没有任何进展时自动终止当前子进程并写入 `inactivity_timeout`，不能让单个阶段无输出持续数小时。
12. **脚本阶段和完整流程必须分开标记**：`scripts\\run_workflow.py` 完成关键词、类目、市场分析和评论抓取后，写入 `scripted_stages_finished` 与 `workflow_handoff_required`；主智能体必须立即调用 `modules\\competitor_analysis\\run.py`，该命令会继续写入同一组监控文件并记录每个后处理阶段的耗时。只有专用竞品分析子智能体完成并通过最终验收后，才允许写入 `workflow_finished`。
13. **监控摘要查看**：使用 `python scripts\\inspect_workflow_monitor.py <缓存目录>\\<name>_workflow_monitor.jsonl` 查看告警、失败任务、已完成评论任务和最后事件。出现告警时，先检查告警前最后一个 `review_more_click_started`、`review_page_ready` 或 `command_started`，再判断是评论页等待、CDP/浏览器连接还是 Excel 后处理。
14. **详细运行诊断**：需要定位耗时或卡点时，运行 `python scripts\\monitor_workflow.py <缓存目录>\\<name>_workflow_monitor.jsonl`，它会按命令列出耗时、输出行数、文件变化、子进程心跳、最长无进展时间和失败原因。运行中的任务可用 `--follow --cdp 127.0.0.1:9223` 实时查看子进程是否存活、最近真实活动时间和 CDP 页面数量；`command_heartbeat` 不代表业务进展，必须结合 `command_output`、`file_activity` 和页面事件判断。

这套限制约束的是无进展等待和重复执行，不是要求在页面尚有明确加载进展时强行截断。有效页面变化、评论数量增加、下载状态更新或输出文件持续生成，都属于可继续监管的真实进展。

### 卡点诊断记录与执行模板

每次用户要求运行并持续监控时，按以下顺序执行，不能先做没有超时保护的手工 DOM 登录态诊断：

1. **先确认是否真的启动**：检查 `workflow_started`、`step_dispatched` 和 `command_started`。如果只有浏览器进程、没有 Python 工作流进程，耗时应归类为“前置环境或诊断阶段”，不能说成业务模块已经运行。
2. **先验证运行环境**：优先使用用户机器上已有的 Python，并通过 `--python` 显式传给工作流；当前机器若没有全局 `python` 命令，可使用 Codex 已提供的 Python 运行时路径。不要因为命令行别名不可用就新建 `.venv`，也不要把 WindowsApps 的 `python.exe` 占位别名当成可用解释器。
3. **正式启动后立即监控**：使用 `scripts\monitor_workflow.py <jsonl> --follow --cdp 127.0.0.1:9223`，同时保留工作流进程的会话。监控输出中的 `command_heartbeat` 只是心跳，不算业务进展；必须结合 `command_output`、`file_activity`、页面数量变化和完成数量判断。
4. **本次已验证的卡点**：2026-09-30 德国站 `DE`、邮编 `10115` 的三模块运行中，正式工作流在 `keyword_backfill.generate` 启动后约 60 秒开始无进展告警；约 2 分钟内没有任何子进程输出、文件变化或后续步骤事件，CDP 仍只有 SellerSprite 和 Amazon 两个页面。该次应归类为“关键词导出浏览器子进程卡住”，不是类目截图或市场分析耗时，已按规则停止，未继续盲等。
5. **浏览器诊断边界**：不要在工作流外重复调用无超时的 Playwright `page.evaluate()` 检查登录态；遇到页面渲染器无响应时，这类诊断本身可能无限等待。优先读取监控日志、进程状态和 CDP 标签页数量；需要修复时，应在脚本内为页面探测和 CDP 连接增加明确超时，并只重启本项目浏览器一次。
6. **停止后保留证据**：停止任务后保留对应缓存目录中的 `{name}_workflow_monitor.log`、`{name}_workflow_monitor.jsonl`、配置文件和工作目录，先根据最后一个 `command_started`、最后一个真实输出和 `inactivity_warning` 定位，再决定从检查点续跑或修改脚本。不得删除日志后重新猜测卡点。

标准启动示例（只运行市场调研与市场分析，站点和邮编必须显式指定）：

```powershell
& "<已有Python>" ecommerce-workflow-skill\scripts\run_workflow.py `
  --source "C:\路径\输入.xlsx" `
  --site DE `
  --amazon-postal-code 10115 `
  --steps keyword_research,category_research,market_analysis `
  --category-overviews `
  --category-overviews-debugger-address 127.0.0.1:9223 `
  --keyword-backfill-debugger-address 127.0.0.1:9223 `
  --keyword-trends-debug-port 9223 `
  --market-analysis-charts `
  --output-dir "C:\路径\缓存目录" `
  --final-output "$HOME\Desktop\产品名-德国站-市场调研与市场分析.xlsx" `
  --python "<已有Python>" `
  --node "<已有Node>"
```

## 站点规则

当前流程支持用统一参数 `--site` 控制卖家精灵站点和 Amazon 站点。默认是 `AUTO`：数据源只有一个 Amazon 域名时自动识别；没有域名时兼容回退为 `US`；检测到多个站点时必须明确指定。

```text
US -> 卖家精灵美国站 -> amazon.com -> 默认邮编 10001
DE -> 卖家精灵德国站 -> amazon.de -> 默认邮编 10115
UK -> 卖家精灵英国站 -> amazon.co.uk -> 默认邮编 B100AB
CA -> 卖家精灵加拿大站 -> amazon.ca -> 默认邮编 M4Y1M7
FR -> 卖家精灵法国站 -> amazon.fr -> 默认邮编 75000
```

执行规则：

- `keyword_research` 凡是进入卖家精灵页面，都必须先确认卖家精灵登录态，再确认当前站点是否与 `--site` 一致；不一致时先切换站点。
- `keyword_research` 抓 Amazon 站内结果数时，必须访问目标站点域名，并先确认 Amazon 登录态和目标邮编。
- `category_research` 抓的是 Amazon 页面里的 softtime/sorftime 插件区域，不检查卖家精灵登录态，但必须确认 Amazon 类目 URL 域名与 `--site` 一致，并确认目标邮编后再截图。
- `review_research` 直接抓目标 Amazon 站点的商品评论页，不依赖卖家精灵或 softtime/sorftime；开始抓取前必须确认 Amazon 登录态和目标邮编，出现登录页或验证码时停止并明确报告。
- Amazon 邮编确认必须以页面左上角配送位置文本为准；不能把 `unknown` 当成功。英国站输入 `B100AB` 后，Amazon 顶部可能显示为 `Birmingham B100AB`，这属于成功；但如果顶部仍显示 `China` 或不包含目标邮编/邮编片段，必须停止相关截图和站内结果抓取。
- 德国站流程需要数据源里的类目超链接本身就是 `amazon.de`；英国站流程需要数据源里的类目超链接本身就是 `amazon.co.uk`；加拿大站流程需要数据源里的类目超链接本身就是 `amazon.ca`；法国站流程需要数据源里的类目超链接本身就是 `amazon.fr`。如果目标站点和 Excel 里的 Amazon 域名不匹配，脚本应停止并提示域名不匹配，避免跨站点截图造成数据错误。
- 旧参数 `--keyword-backfill-site` 仅作为兼容别名保留；新流程优先使用 `--site`。

## 项目目录

```text
亚马逊选品自动化SKILL\
  ecommerce-workflow-skill\
    SKILL.md
    .gitignore
    agents\
    docs\
    scripts\
      setup_first_run.ps1
      run_workflow.py
      open_sellersprite_login_browser.ps1
      convert_wps_dispimg_to_floating.py
      make_wps_embedded_version.py
    modules\
      keyword_backfill\
      keyword_trends\
      product_research\
      market_analysis\
      review_research\
      competitor_analysis\
      shared\
```

模块职责：

- `keyword_research`：关键词调研，对外是一个步骤；内部先打开卖家精灵下载关键词拓展 Excel、筛选并回填关键词，再抓取关键词趋势截图、趋势字段和 Amazon 站内结果数。源码仍分布在 `modules\keyword_backfill` 和 `modules\keyword_trends`，旧步骤名保留兼容。
- 关键词趋势字段中，卖家精灵月搜索量为 `1` 通常是趋势图显示异常或占位值，不参与淡季月搜索量、旺季月搜索量和趋势判断基准；不能把 `1` 写成淡季搜索量。
- 关键词在最近完整年没有 12 个有效月搜索量时，视为后收录或数据不足词：从第一个有效月搜索量起取到最新月份计算趋势、淡季月搜索量和旺季月搜索量，不得因上一完整年缺数而写入 `/`。
- 关键词趋势字段里的“旺季竞价”必须按 `--site` 写入对应货币格式：US 为 `$`，DE 为 `€`，UK 为 `£`，CA 为 `CA$`，不能沿用模板里残留的其它站点货币符号。
- `category_research`：抽取数据源记录，填充产品调研模板，生成市场调研第 2 小节分类节点，抓取类目表现截图，并回填类目总月销量。代码仍放在 `product_research` 模块中。
- 类目表 I 列保留数值型月总销量，下一行显示 `产品占比：xx.x%`。按有效类目产品条数计算占比，保留原底色：大于 50% 为浅蓝，其余为浅粉。已有报告只改占比时，可使用 `python -m modules.product_research.scripts.refresh_category_shares --workbook <报告> --source <含数据源的工作簿> --output <独立副本>`，不重抓截图；数据源改变后需重新运行以更新占比。
- `category_research` 的类目表现截图等待目标是 softtime/sorftime 市场全景区域本身，不是整个页面。类目按节点 ID 去重后必须处理全部类目；截图区域出现后固定等待满 20 秒，再检查所需文字和全部有效 canvas 并截图；目标区域出现最多等待 30 秒，固定等待结束后按当前已加载内容截图并继续下一个类目。截图调用本身必须设置短超时，不能因 CDP 截图无限挂起。
- 类目截图不能只依据文字、canvas 数量和宽高判断完成；截图后必须检查像素是否为纯白或近似空白。检测到空白图时，在 30 秒上限内短暂重试，必要时切换 CDP surface 截图兜底；最终仍需保留当前截图并在日志中标记 `screenshot_blank`。
- `market_analysis`：先生成并插入市场调研第 3 小节 9 张 ECharts 图表，再生成市场分析相关 sheet 和最终报告内容。市场分析页的销量分层、评分数区间和价格段 3 张图必须冻结为静态 PNG 后再进入评论和竞品步骤，避免复用模板改 Sheet 名或后续 `openpyxl` 保存后出现空白图表。
- `review_research`：从数据源 ASIN 列或 `--review-asin` 读取商品，分页抓取 Amazon 评论标题、正文、星级、日期、变体、Verified Purchase、Helpful、图片/视频标识等字段，并写入评论摘要和明细 sheet。完整流程必须显式传入“竞品分析”页列出的竞品 ASIN，不能抓取整个数据源。父子体共享评论时始终直接使用当前子 ASIN，由 Amazon 自己路由共享评论池，不替换或回退到父 ASIN。
- `keyword_trends` 默认复用 `127.0.0.1:9223` 的项目专用浏览器，不再在完整流程中默认关闭所有 Chrome 后另起浏览器；只有显式传入 `--no-keyword-trends-use-existing-browser` 时才允许启动独立浏览器。
- `competitor_analysis`：处理竞品分析页。当前已实现评价分析子模块：按竞品 ASIN 随机抽取好评样本供 Agent 总结优点，按每个 ASIN 的归一化差评覆盖率回填缺点，并综合全部竞品生成品类总结和 `缺点 -> 解决办法`。
- `shared`：只放多个模块共用的工具函数。单一业务逻辑不要放进这里。

## 环境要求

推荐环境：

- Windows。
- Python 3.10 或以上。
- Node.js 18 或以上。
- 已安装 Google Chrome 或 Microsoft Edge。
- 卖家精灵账号，并且有导出权限。
- 如需抓取亚马逊站内结果，需要可用的 Amazon 登录态。
- WPS 或 Excel 可用于打开最终 `.xlsx` 文件。

常用 Python 依赖：

- `openpyxl`
- `pandas`
- `numpy`
- `Pillow`
- `playwright`

产品调研图表运行时在：

```text
modules\product_research\scripts\chart-runtime\package.json
```

其中 vendored ECharts 文件必须保留：

```text
modules\product_research\scripts\chart-runtime\vendor\echarts.min.js
```

分发项目时不要删除这个文件。

## 新电脑初始化

进入内层项目目录后执行命令：

```powershell
cd "$HOME\.codex\skills\ecommerce-workflow-skill"
```

首次使用前，优先运行一键准备脚本：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup_first_run.ps1
```

这个脚本会做三件事：

- 在项目目录下创建 Python 虚拟环境：`.venv`。
- 按根目录 `requirements.txt` 把工作流需要的 Python 依赖安装到 `.venv`，不安装到用户电脑的全局 Python。
- 把 Playwright Chromium 安装到项目 `.venv` 相关目录，尽量避免污染用户全局缓存。
- 在 `modules\product_research\scripts\chart-runtime` 安装产品调研图表运行时依赖。
- 运行 `scripts\check_environment.ps1 -RequireInstalled` 做安装后自检。

首次准备完成后，后续运行 workflow 默认使用：

```powershell
.\.venv\Scripts\python.exe
```

`scripts\run_workflow.py` 也内置了保护：如果项目 `.venv\Scripts\python.exe` 已存在，但用户误用全局 `python scripts\run_workflow.py` 启动，它会自动切换到项目 `.venv` 的 Python 后再继续执行。各子模块收到的 `--python` 默认也会指向这个 `.venv` Python。

`.venv` 是每台电脑本地生成的依赖目录，不要打包分发给别人。别人拿到项目后运行 `scripts\setup_first_run.ps1`，会在自己的电脑上生成自己的 `.venv`。

### 环境依赖路径指向

本项目的依赖路径按下面规则固定，排查环境问题时优先看这些位置：

- Python 虚拟环境目录：`ecommerce-workflow-skill\.venv`
- workflow 使用的 Python：`ecommerce-workflow-skill\.venv\Scripts\python.exe`
- Python 包安装位置：`ecommerce-workflow-skill\.venv\Lib\site-packages`
- Playwright Python 包：安装在 `.venv\Lib\site-packages`
- Playwright Chromium：由 `.venv\Scripts\python.exe -m playwright install chromium` 安装，脚本运行时默认设置 `PLAYWRIGHT_BROWSERS_PATH=0`
- 产品调研 ECharts/Node 依赖：`modules\product_research\scripts\chart-runtime\node_modules`
- Node 可执行文件：默认使用系统 `node`，也可以运行时通过 `--node "node路径"` 指定；Node 包不装进 `.venv`

`scripts\run_workflow.py` 启动时会打印实际使用的 `python` 和 `node` 路径。如果日志里的 `python` 不是 `.venv\Scripts\python.exe`，说明当前运行方式不符合标准，应先运行 `scripts\setup_first_run.ps1` 或检查 `.venv` 是否存在。

如果一键脚本不可用，再手动执行下面的命令。

安装 Python 依赖：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
.\.venv\Scripts\python.exe -m pip install -r modules\keyword_backfill\requirements.txt
.\.venv\Scripts\python.exe -m pip install -r modules\market_analysis\requirements.txt
.\.venv\Scripts\python.exe -m pip install openpyxl pandas numpy pillow playwright
$env:PLAYWRIGHT_BROWSERS_PATH="0"
.\.venv\Scripts\python.exe -m playwright install chromium
```

安装产品调研图表依赖：

```powershell
cd modules\product_research\scripts\chart-runtime
pnpm install
cd ..\..\..\..
```

如果没有 `pnpm`，可以使用当前机器可用的 Node 包管理器。关键要求是 `render_section3_echarts.mjs` 能正常加载本目录 `node_modules` 里的 `playwright`。

Codex Desktop 注意事项：如果系统 `python` 不在 PATH 中，可以用工作区依赖工具提供的 bundled Python 来执行 `scripts\setup_first_run.ps1 -Python "python路径"` 创建 `.venv`。创建完成后，后续 workflow 都应使用 `.venv\Scripts\python.exe`。不要只依赖 Codex bundled `node_modules`，因为其中可能缺少 `playwright-core`，导致图表渲染报错。最稳妥的方式仍然是在 `chart-runtime` 目录执行 `pnpm install` 或让 `setup_first_run.ps1` 自动安装。

如果确实需要临时指定外部 Node 依赖，`CODEX_NODE_MODULES` 应指向“包含 `node_modules` 的上级目录”，而不是 `node_modules` 目录本身。例如：

```powershell
$env:CODEX_NODE_MODULES="C:\path\to\node-runtime-root"
```

## 首次网页登录

网页登录态只保存在本机项目目录中，不能打包给别人。

运行网页相关步骤前，先检查是否存在项目登录态：

```text
.workflow-state\sellersprite\.browser-state.json
.workflow-state\sellersprite\.chrome-profile\
```

如果登录态不存在或已经失效，打开项目专用 Chrome：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\open_sellersprite_login_browser.ps1
```

这个脚本会优先启动 Chrome，未找到 Chrome 时启动 Edge：

```text
DevTools 端口：9223
浏览器配置目录：.workflow-state\sellersprite\.chrome-profile
默认页面：https://www.sellersprite.com/v3/keyword-miner/
```

首次使用时，必须先弹出这个项目专用 Chrome，再让用户在这个 Chrome 里安装浏览器拓展并登录卖家精灵和 Amazon，方便后续复用插件、站点和登录态。不要把用户名、密码、Cookie 或浏览器状态打包进项目。

### 首次安装浏览器拓展

浏览器拓展必须安装在项目专用 Chrome 配置里，也就是：

```text
.workflow-state\sellersprite\.chrome-profile
```

不要让用户在普通 Chrome、Edge 或其它随机浏览器里安装插件；普通浏览器里的插件不会自动出现在项目专用 Chrome 里，后续自动化也读取不到。

首次安装步骤：

1. 运行 `scripts\open_sellersprite_login_browser.ps1`，弹出项目专用 Chrome。
2. 在这个项目专用 Chrome 里安装流程需要的浏览器拓展，至少包括 Amazon 类目页截图依赖的 softtime/sorftime 插件。
3. 安装后，在同一个项目专用 Chrome 里打开一个 Amazon 类目页，确认页面能出现插件分析区域，或页面 DOM 中能加载 `.fullPanoramaBox`。
4. 继续在同一个项目专用 Chrome 里登录卖家精灵。
5. 继续在同一个项目专用 Chrome 里登录 Amazon。
6. 登录和插件安装完成后，不要切换到普通浏览器；后续 workflow 都复用这个项目浏览器和 `.workflow-state`。

如果首次用户还没有安装 softtime/sorftime 插件，`category_research` 不能继续截图，应先停止并提示用户在项目专用 Chrome 中安装插件。插件安装完成后再重启或继续运行相关步骤。

推荐首次使用顺序：

1. 运行 `scripts\setup_first_run.ps1` 安装依赖。
2. 运行 `scripts\open_sellersprite_login_browser.ps1` 打开项目专用 Chrome。
3. 在这个项目专用 Chrome 里安装 softtime/sorftime 等浏览器拓展，并打开 Amazon 类目页确认插件区域可加载。
4. 在这个项目专用 Chrome 窗口里登录卖家精灵。
5. 在同一个项目专用 Chrome 窗口里打开并登录 Amazon。
6. 登录完成后不要切换到普通浏览器，后续 workflow 复用这个项目浏览器和 `.workflow-state`。

Agent 执行规则：

1. 运行网页步骤前，必须先检查并复用 `.workflow-state`。
2. 不要在检查保存状态之前就要求用户重新登录。
3. 不要使用随机的普通浏览器配置目录。
4. 首次用户未安装 softtime/sorftime 插件时，先运行 `scripts\open_sellersprite_login_browser.ps1` 弹出项目专用 Chrome，并提示用户在这个浏览器里安装插件，不要继续类目截图。
5. 凡是执行 `keyword_research`，都必须在打开卖家精灵工具页后执行登录态守卫。兼容旧步骤 `keyword_backfill` 或 `keyword_trends` 单独运行时也不能跳过。
6. 项目浏览器里已有卖家精灵已登录工具页时，优先点击页面内可见的工具导航切换页面，不要直接刷新或新开目标工具 URL；卖家精灵的会话状态可能依赖当前标签页，直接访问会错误显示为游客。
7. 卖家精灵如果显示 `未登录` 或 `游客`，不要立即判定登录态失效。先点击右上角 `未登录/游客` 打开登录抽屉，再点击 `立即登录`、`登录 / 注册`、`登录`、`登陆`、`Sign in` 或 `Login`。如果页面自动刷新到 `https://www.sellersprite.com/v2/welcome` 或恢复到工具页，并且右上角不再显示 `未登录/游客`，说明登录态仍可复用。
8. 点击抽屉登录后，必须回到真实工具页再复查登录态。不要只因为 `v2/welcome` 能打开就认为后续工具页已登录。
9. 如果卖家精灵或 Amazon 出现密码、验证码、安全验证等页面，停止自动化并让用户手动处理。不要绕过安全验证。
10. `competitor_analysis` 也必须执行同样的登录态恢复规则：项目浏览器未启动时先运行 `scripts\open_sellersprite_login_browser.ps1`；卖家精灵先点击未登录抽屉和登录按钮再复查；Amazon 先通过目标站点 `/gp/css/homepage.html`、站点首页和可见登录入口尝试恢复。不得在自动恢复前直接要求用户登录。

Amazon 邮编规则：

- 默认使用美国邮编 `10001`。
- 如果 Amazon 配送地址变化，点击左上角配送地址，输入 `10001`，点击 `Apply`，再点击 `Done`。
- 中文界面按相同位置或等价按钮处理。
- 抓取 `站内结果呈现` 前，必须先检查 Amazon 是否仍然登录。不要只设置邮编就直接搜索。
- 如果检测到 `Hello, sign in`、登录页、密码框、登录按钮等未登录状态，停止流程并让用户在项目 Chrome 里重新登录 Amazon。不要在未登录状态下继续解析站内结果数。
- 如果出现 `Check the box to continue / I am human` 真人门禁，不得由脚本代勾。运行 `python scripts\amazon_human_gate.py --debugger-address 127.0.0.1:9223`，脚本会把项目 Chrome 置前并等待用户本人勾选；检测到复选框已勾选后再自动点击 `Continue shopping`。完成门禁后从最近检查点继续，不重跑前面已成功的阶段。

如果 Playwright 连接 `127.0.0.1:9223` 超时，但浏览器端口看起来存在，通常是项目 Chrome 进程卡住。处理方式：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\open_sellersprite_login_browser.ps1
```

这个脚本会关闭占用当前端口的项目调试浏览器并重新打开同一个 `.workflow-state\sellersprite\.chrome-profile`，不会新建随机登录态。

## 完整流程运行方式

从内层目录 `ecommerce-workflow-skill` 执行：

```powershell
$env:PYTHONUTF8="1"
$env:PYTHONIOENCODING="utf-8"
python scripts\run_workflow.py `
  --source "C:\path\source.xlsx" `
  --keyword-backfill-debugger-address "127.0.0.1:9223" `
  --keyword-trends-debug-port 9223 `
  --keyword-trends-use-existing-browser
```

`--name` 是可选参数。如果不传，流程会从以 `市场调研` 结尾的 sheet 名里推断产品名称。例如：

```text
Towing Winch Covers市场调研 -> Towing Winch Covers
```

日期后缀优先从源文件名里的 6 位日期后缀读取，例如 `260630`；如果源文件名没有日期，则使用当天 `yyMMdd`。

完整流程最终文件输出到桌面文件夹：

```text
Desktop\亚马逊产品调研流程输出表\产品名称-产品调研-日期.xlsx
```

示例：

```text
Desktop\亚马逊产品调研流程输出表\Towing Winch Covers-产品调研-260630.xlsx
```

Agent 执行规则：

- 用户要求“运行某一步”或“试跑”时，默认不要传 `--final-output`。
- 除非用户明确指定输出路径，否则不要把最终交付 Excel 写到项目目录或 `outputs` 目录。
- 完整流程最终交付文件必须落到桌面 `亚马逊产品调研流程输出表` 文件夹。
- 只要不是完整流程运行完，都视为部分流程；部分流程最终交付文件必须直接落到桌面根目录。
- 部分流程命名规则固定为：`产品名称-步骤名称.xlsx`。不要包含 `产品调研` 和日期后缀，除非用户明确指定。
- `--output-dir` 只是内部缓存目录，不是用户最终交付目录；没有特别原因不要手动传。

## 只运行部分步骤

使用 `--steps` 指定步骤。

只运行关键词调研：

```powershell
python scripts\run_workflow.py `
  --source "C:\path\source.xlsx" `
  --steps keyword_research `
  --keyword-backfill-debugger-address "127.0.0.1:9223" `
  --keyword-trends-debug-port 9223 `
  --keyword-trends-use-existing-browser
```

兼容用法：如果只为排错或维护，也可以继续单独运行旧步骤 `keyword_backfill` 或 `keyword_trends`。业务上默认把它们视为 `keyword_research` 的内部子步骤，不再作为两个独立功能对外介绍。

旧步骤连续运行仍可用：

```powershell
python scripts\run_workflow.py `
  --source "C:\path\source.xlsx" `
  --steps keyword_backfill,keyword_trends `
  --keyword-backfill-debugger-address "127.0.0.1:9223" `
  --keyword-trends-debug-port 9223 `
  --keyword-trends-use-existing-browser
```

如果只运行部分步骤，并且没有传 `--final-output`，面向用户的结果会复制到桌面根目录。命名规则是 `产品名称-步骤名称.xlsx`：

```text
Desktop\产品名称-步骤名称.xlsx
```

示例：

```text
Desktop\Towing Winch Covers-关键词调研.xlsx
Desktop\Towing Winch Covers-关键词导出、关键词趋势.xlsx
```

## 输出和缓存位置

面向用户的文件：

- 完整流程：`Desktop\亚马逊产品调研流程输出表\产品名称-产品调研-日期.xlsx`
- 部分流程：`Desktop\产品名称-步骤名称.xlsx`

内部缓存和中间文件：

```text
Desktop\亚马逊产品调研流程缓存\
  {name}_wps_images_floating.xlsx
  {name}_with_keyword_backfill.xlsx
  {name}_with_keyword_trends.xlsx
  {name}_source_records.json
  {name}_report_filled.xlsx
  {name}_chart_images\
  {name}_report_with_charts.xlsx
  work\{name}\keyword_backfill\downloads\
```

桌面 `亚马逊产品调研流程缓存` 是可删除缓存目录，不是用户最终交付文件。分发项目前可以清空项目内旧 `outputs`，不要把缓存、下载导出表或运行结果打包给别人。

如需改变缓存目录：

```powershell
python scripts\run_workflow.py `
  --source "C:\path\source.xlsx" `
  --output-dir "D:\workflow-cache"
```

只有用户明确需要指定最终输出路径时，才使用 `--final-output`。Agent 不要为了方便调试把 `--final-output` 指到项目 `outputs`。

## 使用已有卖家精灵导出文件

如果卖家精灵关键词拓展表已经手动下载好，可以跳过浏览器下载：

```powershell
python scripts\run_workflow.py `
  --source "C:\path\source.xlsx" `
  --steps keyword_research `
  --keyword-backfill-skip-browser `
  --keyword-backfill-export-file "C:\path\ExpandKeywords-US-xxx.xlsx"
```

这会用指定导出文件完成关键词导出回填，然后继续执行关键词趋势回填。若只想处理指定导出文件、不抓趋势，可使用兼容旧步骤 `--steps keyword_backfill`。

## WPS 嵌入单元格图片

流程默认开启：

```text
--preserve-wps-images true
```

作用是先把 WPS 私有 `DISPIMG` 嵌入单元格图片转换成标准浮动图片，再交给 `openpyxl` 处理，避免产品图片和截图在中间保存时丢失。

可单独运行的工具：

```powershell
python scripts\convert_wps_dispimg_to_floating.py --input "source.xlsx" --output "floating.xlsx"
python scripts\make_wps_embedded_version.py --original "source.xlsx" --floating-workbook "floating.xlsx" --output "embedded.xlsx"
```

不要在完整流程中途把关键词趋势截图重新转换回 WPS `DISPIMG`。所有 `openpyxl` 步骤完成前，应保持图片为浮动图片。

## 数据源 Sheet 命名规则

最终工作簿中的业务数据源页应规范命名为：

```text
产品名称-数据源
```

示例：

```text
cold plunge cover-数据源
```

输入文件可以有两种情况：

- 已经有 `产品名称-数据源`，并且数据就在这个 sheet 里：直接使用它。
- 只有固定名 `数据源`，并且数据就在这个 sheet 里：流程会自动把它改名为 `产品名称-数据源`。

脚本读取数据源时的优先级是：

```text
产品名称-数据源 -> 数据源 -> US -> Sheet1
```

如果同时存在 `产品名称-数据源` 和固定名 `数据源`，后续流程会优先使用 `产品名称-数据源`，并在生成报告时删除多余的固定名 `数据源`，避免最终文件出现两张数据源表。

注意：`产品调研`、`市场分析` 都依赖同一个数据源页。不要让不同模块各自创建独立数据源 sheet。

## 模块说明

### keyword_research

对外入口：

```text
scripts\run_workflow.py --steps keyword_research
```

内部子步骤入口：

```text
modules\keyword_backfill\run.py
modules\keyword_trends\run.py
```

内部主脚本：

```text
modules\keyword_backfill\seller_sprite_keyword_tool.py
modules\keyword_trends\excel-sellersprite-trends.py
modules\keyword_trends\sellersprite_playwright_runner.py
```

输出：

```text
Desktop\亚马逊产品调研流程缓存\{name}_with_keyword_backfill.xlsx
Desktop\亚马逊产品调研流程缓存\{name}_with_keyword_trends.xlsx
Desktop\亚马逊产品调研流程缓存\work\{name}\keyword_backfill\downloads\
```

写入工作簿的卖家精灵导出数据 sheet 固定命名为：`关键词导出`。

规则：

- `keyword_research` 内部先执行关键词导出回填，再执行关键词趋势回填。
- 本步骤凡是操作卖家精灵页面，都要先执行卖家精灵登录态守卫。
- 如果页面右上角显示 `未登录/游客`，先点开登录抽屉并点击登录按钮尝试恢复已保存登录态。
- 只有遇到密码框、验证码或安全验证时，才提示用户手动登录。
- 写回市场调研页的关键词必须显示为蓝色、带单下划线，并带默认 Amazon 搜索超链接，视觉和类目节点链接保持一致。
- 从市场调研 sheet 读取关键词。
- 本步骤要按 `--site` 切换卖家精灵站点；美国站对应 `amazon.com`，德国站对应 `amazon.de`，英国站对应 `amazon.co.uk`。
- 如果关键词文本里包含括号翻译，搜索时只输入真实关键词，不输入括号里的翻译。
- 趋势截图以浮动图片插入。
- 趋势截图导出前必须把卖家精灵 ECharts 强制重排为固定尺寸，最终 PNG 固定为 `3072 x 840`；导出后必须校验像素尺寸，尺寸不符时重新重排并导出，不能把抽屉半宽状态下生成的短图插入 Excel。
- 不要修改模板既有行高和列宽。
- `站内结果呈现` 解析成功时应写数字，解析失败时才写 `/`。
- 抓取 `站内结果呈现` 时，先检查 Amazon 登录态，再确认目标站点邮编，再搜索关键词。Amazon 登录失效时要停止并提示用户重新登录，避免结果数偏差。
- `站内结果呈现` 抓取成功后，需要把脚本实际访问的 Amazon 搜索结果 URL 写入对应关键词单元格的超链接，方便用户点击关键词直接打开站内搜索页。
- 写入或刷新关键词超链接时，必须保持关键词单元格为蓝色、单下划线的链接样式。
- `旺季竞价` 默认对所有关键词使用最新一周 PPC 的“最大值”序列，与是否有旺季无关。先按日期确定最后一周，再读取该周有效竞价；最新周为空或为 `0` 时写 `/`，不回退旧周、近30天最大值或月度数据。`--keyword-ppc-bid-mode peak-period` 可选原旺季区间逻辑。
- 完成关键词回填后按 `旺季月搜索量` 数值降序排列整行；相同值保持原顺序，缺失值排最后。关键词、链接、图片、行样式和字段一起移动，其它分区不变。
- 已完成的报告只需刷新竞价时，使用 `modules/keyword_trends/refresh_ppc_bids.py --source <原表> --output <独立输出> --cache <本次缓存JSON> --site US`。该入口默认只查询最新周 PPC、保留原有趋势与搜索量，再排序并检查关键词结果，不重抓 Amazon 结果数或趋势图。缓存必须同时匹配站点和周期。

常用参数：

```text
--site AUTO|US|DE|UK|CA|FR
--amazon-postal-code 10115
--keyword-trends-debug-port 9223
--keyword-trends-use-existing-browser
--keyword-trends-skip-screenshot
--keyword-trends-keep-artifacts
```

旺季判断规则记录在：

```text
docs\keyword_trend_seasonality_rules.md
```

### category_research

入口：

```text
modules\product_research\run.py
```

执行顺序：

```text
extract_source_records.py
fill_custom_template.py
capture_category_overviews.py
verify_custom_template.py
backfill_category_translations.py (agent post-processing when translations are missing)
```

其中 `fill_custom_template.py` 会优先按数据源类目路径超链接里的 Amazon 末级节点 ID 去重；取不到节点 ID 时，再按规范化后的完整类目路径去重。若同一节点来自多个大类入口，优先采用与类目路径首级一致的大类入口生成链接。脚本据此生成市场调研页第 2 小节，并把最细类目链接写到对应类目词上；链接必须是该类目的 Amazon Best Sellers（BSR）榜单，例如 `https://www.amazon.com/gp/bestsellers/home-garden/723480011`，不能使用普通 `/b?node=` 类目节点页。类目节点要像关键词回填一样显示为 `英文类目` + 下一行 `(中文翻译)`；脚本会优先读取数据源翻译列，没有翻译列时读取 `resources\category_translations.json` 本地翻译缓存。详细规则记录在：

```text
docs\market_research_category_section_rules.md
```

输出：

```text
Desktop\亚马逊产品调研流程缓存\{name}_source_records.json
Desktop\亚马逊产品调研流程缓存\{name}_report_filled.xlsx
Desktop\亚马逊产品调研流程缓存\{name}_category_overviews\
Desktop\亚马逊产品调研流程缓存\{name}_report_with_category_overviews.xlsx
Desktop\亚马逊产品调研流程缓存\{name}_missing_category_translations.json
```

Agent 执行规则：

- `category_research` 可以单独运行。单独运行时也必须先保护 WPS `DISPIMG` 图片，避免产品图、关键词趋势图或已有截图丢失。
- `category_research` 默认不调用翻译 API，也不因新类目的中文缺失而中断。`fill_custom_template.py` 会先按数据源翻译列和 `resources\category_translations.json` 写入已知中文，同时输出 `{name}_missing_category_translations.json`。
- 类目截图、月总销量回填和后续市场分析完成后，agent 必须读取该缺失清单；对缺失英文类目层级自行翻译，生成 `{"English category": "中文翻译"}` JSON，并执行 `backfill_category_translations.py --translations <json> --workbook <最终xlsx>`。该工具会更新本地缓存，并直接修改最终 `.xlsx` 的对应 sheet XML，只刷新类目节点文字，不重新打开网页、不重截图、不删除现有图片。
- 不要在脚本运行前预先要求用户补齐类目翻译；也不要因翻译缺失重跑 `category_research`。目标文件被 Excel/WPS 打开而无法替换时，才请用户关闭该文件后重试翻译回填。
- `category_research` 抓的是 Amazon 类目页上的 softtime/sorftime 插件区域，不是卖家精灵后台页面；不要检查卖家精灵账号登录态。
- `category_research` 凡是要抓类目表现截图，都必须复用项目 Chrome，并确认 Amazon 类目页可打开、目标站点邮编已确认、softtime/sorftime 插件已安装且 `.fullPanoramaBox` 区域可加载。
- `category_research` 必须校验 Excel 类目超链接域名与 `--site` 匹配；例如 `--site DE` 时不能继续抓 `amazon.com` 类目链接。
- 类目表现截图和月总销量回填不要和 9 张 ECharts 图表生成混在同一步里。
- `product_research` 仅作为旧名称兼容，后续 agent 应优先使用 `category_research`。

### market_analysis

入口：

```text
modules\market_analysis\run.py
```

主脚本：

```text
modules\market_analysis\excel_market_report.py
```

该模块使用当前流程工作簿。默认会先生成并插入市场调研第 3 小节 9 张 ECharts 图表，再生成市场分析 sheet 和异常样本 sheet。优先读取 `产品名称-数据源`，其次读取 `数据源`、`US`、`Sheet1`。如果只存在固定名 `数据源`，会自动改名为 `产品名称-数据源`；如果同时存在两者，最终保存前会删除多余固定名 `数据源`。

Agent 执行规则：

- 9 张 ECharts 图表归到 `market_analysis` 前置处理，不归到 `category_research`。
- 单独运行 `market_analysis` 时，默认也要先生成并插入 9 张图表；只有用户明确要求跳过时，才使用 `--no-market-analysis-charts`。
- 插入 9 张图表时仍然要保留工作簿已有图片，不要清掉类目表现截图、趋势图或产品图。
- `report_with_charts.xlsx` 只是 9 张图表的前置检查点，不能直接作为后续最终工作簿；必须继续经过 `excel_market_report.py` 写入市场分析聚合值，再冻结市场分析页的 3 张图表。
- 市场分析阶段结束时必须检查关键聚合单元格（样本数、价格、销量、销售额）非空，并确认 3 张冻结图表都存在；检查失败时立即停止，禁止把空白市场分析页交给评论或竞品模块。
- 市场分析和第 3 小节 9 张 ECharts 图表读取价格、销售额时，必须按数据源表头自动识别货币字段，例如 `价格($)`、`价格(€)`、`月销售额($)`、`月销售额(€)`；输出价格段、价格表头、月销售额环形图标题要沿用识别到的货币符号，不能固定写 `$`。
- 市场分析价格段使用动态分组：先按实际价格相邻差距识别自然断层，价格连续时按四分位距计算自适应步长；不再固定使用 `0-20`、`20-40`、`100+`，也不能因为某个价格段样本少而合并。市场分析页和价格切片器必须复用同一组动态价格段标签。
- 评分数分析不能以月销量是否存在作为纳入条件。评分数有效的产品即使月销量为空，也必须进入对应评分数区间；该组销量按 0 汇总，并在评分数分析图中同时标出样本数，避免高评分数产品被误删。
- 数据源分层切片器是待选功能，默认关闭。普通完整流程或单独运行 `market_analysis` 时保持旧版输出，不增加辅助列、表格和切片器；只有用户明确要求时才传入 `--market-analysis-source-slicers`。
- 启用后新增独立的 `切片联动` sheet，原数据源 sheet 不增加辅助列、不转换表格、不放置切片器。联动页顶部放 `销量分层`、`评分数区间`、`价格段` 3 个切片器，下方从第 12 行开始复制完整数据源表格。
- 联动页复制的数据末尾写入同名的 3 个辅助列并隐藏；分组规则必须复用市场分析计算函数。没有有效月销量、因此未进入市场分析的记录，3 列都写 `未纳入市场分析`。
- 联动页的产品图片必须转换成 WPS `DISPIMG` 单元格内图片，不能只设置浮动图片的 Placement。切片器筛选数据行时，对应图片应随单元格一起隐藏或显示。
- 启用后执行 `add_source_slicers.ps1`，只把 `切片联动` 页第 12 行开始的数据转换成可筛选表格。切片器可组合使用，筛选结果取交集；脚本必须幂等，重跑时更新本功能自己的表格和切片器，不能重复堆叠，也不能删除其他切片器。
- 创建联动页切片器依赖本机 Excel/WPS COM。执行顺序必须是：先由 `openpyxl` 生成联动页和浮动图片副本，再把联动页图片转换成 WPS 单元格图片，最后创建原生切片器；创建切片器后不能再由 `openpyxl` 保存该工作簿。

启用示例：

```powershell
python scripts\run_workflow.py --source "C:\路径\产品调研.xlsx" --steps market_analysis --market-analysis-source-slicers
```

### review_research

入口：

```text
modules\review_research\run.py
modules\review_research\scrape_amazon_reviews.py
modules\review_research\backfill_review_translations.py
modules\review_research\backfill_pain_phrase_groups.py
```

单独运行示例：

```powershell
python scripts\run_workflow.py `
  --source "C:\Users\用户名\Desktop\产品调研.xlsx" `
  --steps review_research `
  --site US `
  --review-asin-limit 3 `
  --review-max-pages 5 `
  --review-max-positive-reviews 50 `
  --review-max-critical-reviews 100
```

只抓指定 ASIN：

```powershell
python scripts\run_workflow.py `
  --source "C:\Users\用户名\Desktop\产品调研.xlsx" `
  --steps review_research `
  --site US `
  --review-asin B0XXXXXXXX `
  --review-max-pages 3
```

写入工作簿的 sheet：

```text
评论抓取摘要
商品评论
差评痛点
```

Agent 执行规则：

- 没有传 `--review-asin` 时，按 `竞品ASIN -> 商品ASIN/ASIN -> 子ASIN -> 父ASIN` 的优先级选择数据源列并去重，不把父子 ASIN 两列混在一起；传入一个或多个 `--review-asin` 时，只抓明确指定的 ASIN。
- 单独运行时，沿用局部步骤输出规则，最终文件直接生成到桌面，命名为 `产品名称-评论调研.xlsx`；缓存中间文件仍放在桌面 `亚马逊产品调研流程缓存`。
- 默认只设置 `Top reviews` 排序，并分别加载 `Positive reviews`（4～5 星）和 `Critical reviews`（1～3 星）两组，不主动选择 `All reviewers` 或 `Verified Purchase only`。对应 URL 只传 `sortBy=helpful` 和 `filterByStar=positive/critical`，不传 `reviewerType`。页面评论加载后再做可信度筛选：只保留 `Verified Purchase` 或 Vine 绿标至少满足一项的评论；两项都不满足的评论不写入 Excel。不能要求两项同时满足，因为 Vine 试用评论通常不是普通真实购买。每条保留的明细仍记录 Verified Purchase 和绿标状态。脚本还必须读取页面的实际排序和好评/差评筛选值进行复核，不能只假设 URL 参数已经生效。
- 默认每个 ASIN 最多读取 Positive reviews 候选评论 50 条、Critical reviews 候选评论 100 条，两组候选池最多共 150 条；达到候选数量上限后立即停止，不得为了凑满可信评论数量继续加载。候选评论再按 `Verified Purchase` 或 Vine 绿标过滤，因此最终用于分析并写入 Excel 的数量可以少于 50/100。每组最多允许加载 10 批。分别用 `--review-max-positive-reviews` 和 `--review-max-critical-reviews` 调整；旧参数 `--review-max-reviews` 仅作为兼容覆盖项，显式传入时会同时覆盖两组候选上限。`--review-max-pages` 为兼容旧命令保留，实际表示每组最多加载多少个“每批约 10 条”的评论批次。首次验证或用户未说明范围时，先用 `--review-asin-limit` 或显式 ASIN 小范围运行，不能直接对上百个 ASIN 全量抓取。
- `--review-star-filter both` 是默认值；排错时可改为 `positive` 或 `critical` 单独抓一组。`--review-reviewer-type default` 和 `--review-sort-by helpful` 是默认值；`default` 表示不传 `reviewerType`，只有用户明确要求 `all` 或 `verified` 时才调整。
- 整批评论抓取开始前访问目标站点账户主页，完整检查一次 Amazon 登录态并确认目标邮编；不要在每个 ASIN 的 Positive/Critical 任务前重复跳转账户主页。进入每个评论页后仍必须读取页面左上角配送位置并识别登录页或安全挑战；只有显示目标邮编或允许的邮编片段才能解析评论。若评论页恢复为 `China` 或其它位置，必须先重新设置目标邮编、刷新同一评论 URL 并复核，不能直接把当前页面判定为 0 条评论。
- Amazon 新版评论页通过 `[data-hook="show-more-button"]` 的 `Show 10 more reviews` 按钮懒加载下一批评论。必须点击该按钮并等待 `[data-hook="review"]` 数量增加，不能通过递增 URL 的 `pageNumber` 参数模拟翻页。
- Positive reviews 达到 50 条、Critical reviews 达到 100 条时分别正常停止；未达到对应目标但 `Show 10 more reviews` 按钮已经消失时也立即停止，并在摘要中标记 `insufficient_reviews`，写明实际条数。达到批次上限但按钮可能仍存在时标记 `batch_limit_reached`，不能误报为评论总数不足。
- Amazon 评论页出现验证码、机器人验证、重复批次或空评论卡片时，将状态和原因写入 `评论抓取摘要`，不得把失败静默当成 0 条评论。
- Amazon 父体或合并变体评论池可能混入其它商品；抓取前读取当前 ASIN 数据源中的商品标题和产品卖点，对评论标题/正文做异品特征校验。命中灯具、婴儿用品、工具等强异品短语时剔除并记录 `review_filtered_product_mismatch`，无法确定的普通评论继续保留，不能只按评论国家判断归属。
- `评论抓取摘要` 顶部必须同时包含 Positive 和 Critical 两行，除非用户明确要求只抓其中一组。列顺序固定为：抓取时间、站点、竞品ASIN、评论类型、实际筛选、抓取状态、抓取评论数、Verified Purchase 数量、绿标评论数量、含图片评论数、含视频评论数、商品链接、错误/说明。抓取时间列必须足够宽，不能显示为 `########`。摘要不写商品标题、评论来源ASIN、来源说明、加载批次、请求批次上限、平均星级、各星级数量或评论页链接。该摘要表下方空一行后必须放置独立的“差评词频分析”区块，供用户在同一总览 sheet 快速判断；不再创建单独的 `差评词频分析` sheet。
- 写入或刷新评论 sheet 时必须保留源工作簿已有的浮动图片、绘图锚点和其它页内容；若 openpyxl 未能读取 OOXML 浮动图片，先从 xlsx 绘图关系中恢复后再保存，不能因评论抓取删除竞品产品图。
- 摘要和明细都必须标记 `Positive reviews` 或 `Critical reviews`。`商品评论` 明细列顺序固定为：站点、竞品ASIN、实际筛选、评论日期、星级、评论类型、评论标题、评论正文、购买变体、Verified Purchase、绿标评论、Helpful数、图片URL、含视频、评论链接。该 sheet 不写商品标题、评论来源ASIN、评论 ID、日期原文、评论者、绿标原文、加载批次或来源页面；这些原始字段仍可在抓取过程中用于去重和生成其它 sheet。Vine 绿标优先读取 `[data-hook="vine-review-badge"]`，并兼容 `.cr-vine-review-badge`。
- 每次抓取后自动生成 `差评痛点` sheet，只写入 Critical reviews。每行对应一条差评，列顺序固定为：竞品ASIN、实际筛选、星级、评论日期、属性/型号、差评词/短语、归一化痛点（中文）、Helpful数、图片数量、评论图片、评价内容、评论链接。“差评词/短语”保留每条评论最多 4 个原始英文痛点词或短语；右侧“归一化痛点（中文）”必须与原始短语逐行对应，即使多个原始短语归到同一标准痛点也不能在明细页合并或省略。内置规则和缓存尚未识别的项先显示“待归类”，Agent 分组后由 `backfill_pain_phrase_groups.py` 同时回填该列并刷新 `差评词频分析`。竞品 ASIN 用于对应竞品分析列；All variants 评论的实际变体仍以“属性/型号”列为准。表头和全部数据单元格必须带四边细框线，不能仅靠隐藏网格线形成无边界白色区域。
- `差评痛点` 的图片列只允许放嵌入图片，不写 `查看原图`、`/`、图片 URL 或下载状态等任何文字，也不设置超链接；没有图片或图片下载失败时保持空白，是否有图以“图片数量”为准。需要查看原图时由用户点击该行“评论链接”进入 Amazon。一条评论有多图时，把抓到的图片自适应排成最多 3 列的缩略拼图，并全部放在同一个“评论图片”单元格内。拼图按 Excel 实际展示尺寸生成并保持原始宽高比，不能拉伸变形；不要对 Amazon 本身模糊的评论图片进行无效的高清放大，以免只增加工作簿体积。
- 评论逐条翻译只处理 `Critical reviews`。抓取脚本先读取 `resources\review_translations.json` 中已有差评翻译，并把已命中的中文追加在“差评痛点”的英文评价内容下方；只有缓存未命中的差评正文才输出到 `{name}_missing_review_translations.json`。`Positive reviews` 保留英文原文，不进入缺失翻译清单，也不逐条生成中文，以免为竞品优点总结产生不必要的翻译耗时。
- `review_research` 脚本完成后，Agent 必须确认缺失翻译 JSON 的 `review_scope` 为 `Critical reviews only`，只把其中 `translations` 对象的空值翻译为简洁、忠实的中文，再执行 `backfill_review_translations.py --translations <已翻译json> --workbook <最终xlsx>`。差评评价内容采用 `英文原文 + 下一行 (中文翻译)`，并更新本地缓存及评价内容行高；该工具只能修改对应文字和行高，不能重跑评论抓取或删除现有图片。没有缺失项时跳过回填。
- `差评痛点` 保留每条评论实际提取出的原始英文词/短语，便于核对原文；摘要页下方的“差评词频分析”区块则按语义合并成标准痛点，例如 `sticky surface/sticky touch/sticky feeling` 合并为 `surface stickiness / 表面黏腻`，`stopped working/never worked/broke` 合并为 `product failure / 产品失效`。先使用代码内置规则，再读取 `resources\pain_phrase_groups.json` 的历史 Agent 分组缓存。全局缓存只作为候选映射，必须结合当前评论正文做品类上下文校验；飞行、螺旋桨、降落伞、蒸汽、清洁、耳道、摄像头等领域映射在正文没有对应领域词时不得命中，必须回到当前产品的“待归类”清单，不能跨产品复用。
- 内置规则和缓存都未识别的原始短语必须写入 `{name}_missing_pain_phrase_groups.json`，不能阻塞抓取，也不能擅自丢弃。脚本完成后，Agent 逐项填写 `groups` 中的 `group_en` 和 `group_zh`；含义相近的原始短语必须使用完全相同的标准英文和中文名称。然后执行 `backfill_pain_phrase_groups.py --groups <已分组json> --workbook <最终xlsx>`。该工具从 `差评痛点` 重新计算并原位刷新 `差评词频分析`，同时逐行回填“归一化痛点（中文）”并更新缓存；不能重新抓网页，也不能删除工作簿现有图片。
- 每次抓取后在 `评论抓取摘要` 底部自动生成“差评词频分析”区块，列顺序固定为：竞品ASIN、实际筛选、关键词/短语、中文翻译、评论覆盖数、覆盖率、涉及型号。只基于 Critical reviews 统计，并按竞品 ASIN 和实际筛选分开计算；评论覆盖数按标准痛点对应评论的并集计算，同一条评论即使命中多个同义短语也只计 1 次，覆盖率为该组覆盖数除以该组差评总数；涉及型号按覆盖评论去重汇总。该区块不写排名、总出现次数、原始变体词清单、典型评价、评价正文或评论链接，用户需要查看原始短语和原文时回到 `差评痛点` 或 `商品评论` sheet。

### competitor_analysis：评价分析

完整竞品分析不是普通纯脚本步骤。用户要求竞品分析时，主智能体必须派出一个专用子智能体负责语义分析，并把输入工作簿、输出工作簿、站点、邮编和浏览器地址传给它。语义 handoff 完成后，主智能体必须立即调用统一后处理入口，不能停在交接点等待下一轮对话。子智能体必须先读取：

```text
modules\competitor_analysis\workflow.md
modules\competitor_analysis\agent_prompt.md
```

- 用户明确说“竞品分析里的评价分析暂时不用”时，等同于评论分析也不执行：完整流程中必须跳过 `review_research`，不得抓取评论，不得生成“评论抓取摘要”“商品评论”“差评痛点”，只运行市场调研、类目、市场分析和竞品基础字段。不得把“完整流程”理解为必须运行用户已明确排除的评论模块。

主智能体不要和子智能体同时操作同一工作簿或浏览器页面，也不要重复做同一分析；子智能体完成后，主智能体立即执行：

```powershell
python modules\competitor_analysis\run.py `
  --workbook "<竞品基础回填工作簿.xlsx>" `
  --analysis "<已填写的竞品分析 handoff.json>" `
  --output "<最终输出.xlsx>" `
  --monitor-jsonl "<name>_workflow_monitor.jsonl" `
  --monitor-log "<name>_workflow_monitor.log"
```

该入口连续完成回填、差评透视和原生切片器，并复用原监控起始时间；之后主智能体负责核对 ASIN 对齐、图片、超链接、边框、公式和视觉排版。若当前运行环境没有子智能体能力，主智能体才按 `agent_prompt.md` 的同一角色和边界自行执行，并在结果中说明未能实际委派。

主智能体验收时还必须确认：每个目标竞品的 Positive/Critical 摘要与评论明细数量对平；状态只能是 `success`、`insufficient_reviews` 或评论数确实为 0 的 `no_reviews`，其它状态必须先修复或重跑；缺失翻译和“待归类”痛点均为 0。任一条件不满足时不能把竞品分析报告为完成。

`workflow.md` 覆盖本模块全部工作：产品主图、详情链接、评分/评论数、产品卖点、动态对比维度、父体共享评论、优缺点评价、品类总结、`缺点 -> 解决办法` 和差评透视切片器。以下规则是评价分析阶段的关键约束：

- 评价分析必须在 `review_research` 已生成并完成翻译、痛点归一化之后执行。竞品列按竞品分析页“详情页链接”中的 ASIN 对应，不能按评论页当前显示顺序猜测。
- 对每个竞品，从 `Positive reviews` 中稳定随机抽取最多 25 条；不足 25 条时读取全部。好评样本直接保留并读取站点原文，不要求先逐条翻译；Agent 根据英文或目标站点原文直接生成中文优点，综合样本归纳有评论证据支持的结论，不为凑数量添加结论。
- 每条优点不超过 20 个汉字，可以使用完整、自然的短句，不要为了精简而只保留过短标签；回填时每条独占一行，不添加 `1. 2.` 等数字编号。相同含义必须合并，避免“好用/容易使用/使用方便”重复占点。
- 缺点直接采用归一化痛点中文名称和评论覆盖率，内容格式为 `产品失效（50%）`；同样每条独占一行且不添加数字编号。覆盖率按该 ASIN 的差评数量独立计算；多 ASIN 时不能使用混合后的总比例。
- 多竞品分析必须把竞品分析页中的全部目标 ASIN 直接传给 `review_research`，并按摘要、商品评论和差评痛点页的“竞品ASIN”列归属评论。每个竞品始终用当前子 ASIN 的 `product-reviews` URL，默认不传 `reviewerType`，不读取或替换成父 ASIN，也不生成“评论来源ASIN”。当前子 ASIN 没有文字评论时，其优点和缺点填 `/`，不能复用其它竞品的评论。
- `竞品分析` 页默认必须保留 5 个产品栏位（`C:G`）。竞品不足 5 个时，剩余栏位保留模板边框、尺寸、对齐和合并区域但不填任何内容；竞品超过 5 个时，按 `G` 列对应行的格式向右复制，并把标题和总结正文区域扩展到最后一个竞品列，不能只格式化实际数据中的前几列。所有竞品产品列统一使用 `C:G` 的模板列宽，新增列不能继承模板尾部 `H:P` 的窄列宽；标题、粉色填充、边框和合并区域不得只停留在 `G`。
- 竞品分析首行标题必须保留模板富文本和各文字段的字体，以及主题填充、细边框、左对齐、自动换行和行高。标题合并区域未变化时禁止先取消再重建；确需向右扩容时，重合并后必须恢复原标题和首行样式。最终门禁必须将交付文件与原始模板的首行格式逐项比对。
- WPS 单元格图片编号在整个工作簿内必须唯一。嵌入新图片前必须读取原工作簿已有 `cNvPr.name` 并从下一个可用编号继续；完成后必须检查 `cellimages.xml` 无重复名称，并实际渲染竞品页确认每个 `DISPIMG` 单元格显示的是对应 ASIN 的主图。
- 竞品“产品图”必须优先读取数据源中按 ASIN 对应的产品主图，并保持原始像素、仅按比例缩小显示框以适配各自产品图单元格。只有数据源缺少目标 ASIN 主图时，才回退运行 `modules\competitor_analysis\capture_product_cards.mjs` 截取单个商品卡；禁止默认用商品卡截图替代产品主图。写入后恢复产品列细边框，验证 `TopLeftCell` 与 `BottomRightCell` 均未越出目标单元格。
- `竞品分析` 页的对比维度产品值必须把 `；`/`;` 分隔属性改为单元格内换行，并仅把带单位的可比较规格数值（如 `600D`、`900D`、`8000-13500磅`、`80/250英尺`、`3秒`）设置为红色加粗富文本；斜杠多档、区间、重量、距离、时间、次数和数量均需识别，描述文字保持模板原样，空白竞品栏位不得写入内容。
- “产品卖点”中的可比较规格数值也必须按相同规则局部红色加粗，编号、换行和其余文字保持原样，不得误标 `1. 2.` 等列表编号。“总结及改进方式”不得自行重设版式：左侧标题和右侧正文沿用模板字体、颜色、对齐和换行，只把正文合并区域扩展到最后一个竞品列；所有行高保持模板原值，脚本不得根据文本行数自行加高。
- `prepare_review_analysis.py` 只生成待总结 JSON，不修改工作簿；除了好评样本和差评覆盖率，还必须输出每个竞品的评分/评论数、产品卖点和对比维度。Agent 填写各产品的 `advantages`、顶层 `overall_summary` 和 `improvements` 后，再由 `backfill_review_analysis.py` 统一写入。
- `overall_summary` 是综合全部竞品后的品类总结，用 1-2 句概括主要体验、共同优势和产品成熟度，保持精简，不逐个复述 ASIN。
- `improvements` 每项固定包含 `pain_point` 和 `solution`。应把语义相近的缺点合并，优先覆盖高频、共性和安全性问题；解决办法必须具体、可执行，不能只写“提高质量”“优化体验”等空泛内容。Excel 中每项显示为 `缺点 -> 解决办法`，每条独占一行，不添加数字编号。
- 回填脚本只修改有评论数据的竞品优缺点列，并把总体总结和改进方式写入“总结及改进方式”右侧合并区域。没有抓取评论的竞品优缺点保持 `/`。写入后必须恢复对应区域四边细框线，保留模板行高、产品图、详情链接以及其它 sheet。
- 评价分析完成后把“差评痛点”全部明细迁移到“差评透视”并删除原 sheet：按“评论 × 归一化痛点”展开差评。总标题、产品图片/维度筛选分区、明细表头和切片器统一使用橙色体系。版式固定为第 1 行“差评透视”总标题，第 2 行“产品图片 / 维度筛选”分区，第 3 行保持样例的 `409pt` 行高，左侧动态商品卡，右侧在既定列宽内等宽、等间距依次铺开 `ASIN`、`属性/型号`、`归一化痛点（中文）` 3 个橙色原生切片器，不能缩成小控件或重叠，第 4 行开始明细表。切片字段不得显示 `/`：缺少型号写“未注明型号”，没有有效归类写“未归类痛点”。组合条件取交集，选择单个 ASIN 时上方图片和下方差评必须属于同一竞品。商品卡图片库必须保留原始像素高清源，不能先缩小再放大；动态展示仅按比例向下适配图片区并关闭图片压缩。
- 差评评论图片必须先做有效像素检查。下载失败后形成的纯白、近纯白或透明占位 PNG 不得写入“差评透视”，对应图片数量写 `0`；有实际画面的评论图才允许随筛选行显示。
- 先运行 `modules\competitor_analysis\build_pain_slicer_view.py` 生成联动页和隐藏图片库，再运行 `modules\competitor_analysis\add_pain_slicers.ps1` 创建切片器及动态商品图。切片器创建必须是整个竞品分析的最后一个 Excel 写入步骤，之后严禁再用 `openpyxl` 保存该文件。

## 本项目已知卡点和处理方式

这些是实际运行中遇到过的问题，交给别人使用前必须在首次准备阶段处理掉：

- 图表渲染报 `Cannot load playwright`：进入 `modules\product_research\scripts\chart-runtime` 执行 `pnpm install`。不要等跑到 `market_analysis.render_charts` 才处理。
- Codex bundled Node 里可能只有 `playwright` 壳而缺 `playwright-core`：优先使用 chart-runtime 本地 `node_modules`，不要依赖 bundled `node_modules`。
- `pnpm install` 如果提示 `sharp` 构建脚本被忽略，但图表能正常渲染，可以继续；本流程图表生成主要依赖 Playwright 和 ECharts。
- `BrowserType.connect_over_cdp` 连接 `9223` 单次超过 20 秒：先检查 `/json/version`，再用 `scripts\open_sellersprite_login_browser.ps1` 重启同一项目浏览器配置并只重连 1 次；仍失败就停止当前阶段并报告，禁止反复等待 180 秒。
- 运行网页步骤前必须先检查 `.workflow-state`，有登录态就复用；没有或失效才让用户登录卖家精灵和 Amazon。
- 卖家精灵页面显示 `未登录/游客` 时，先点右上角打开抽屉并点击登录按钮恢复保存的登录态。不要直接判定失败，也不要让流程卡在这个判断上。
- `category_research` 是例外：它抓 Amazon 页面里的 softtime/sorftime 插件数据，不进入卖家精灵后台，不需要检查卖家精灵登录态。
- Amazon 登录态可能比卖家精灵更早过期。抓站内结果时如果发现未登录，必须停止让用户重新登录 Amazon，不能继续用未登录搜索结果。
- `review_research` 同样依赖 Amazon 登录态；评论页出现 `/ap/signin` 或验证码时必须停止或记录为受阻，不能输出看似成功的空评论表。
- 完整流程验收不能只数图片或图表对象。必须抽查实际像素或通过 Excel 复制图片验证市场分析图表、类目截图、竞品商品卡和差评评论图确实有内容；发现空白时只从最近的无切片器检查点重建失败模块，不重跑已经完成的浏览器抓取和评论抓取。
- 源数据可以放在 `产品名称-数据源`，也可以只放在固定名 `数据源`。如果只给 `数据源`，流程会自动改成规范名；如果两者同时存在，脚本会优先用 `产品名称-数据源` 并清理多余固定名 `数据源`。

## 分发前检查

把项目发给别人之前：

1. 清空或删除项目内旧 `outputs\`，并删除桌面 `亚马逊产品调研流程缓存` 中不需要的缓存。
2. 删除 `.workflow-state\`。
3. 删除 `.venv\`，不要把本机 Python 依赖环境打包给别人。
4. 删除所有 `__pycache__\` 和 `*.pyc`。
5. 不要包含浏览器配置、Cookie、卖家精灵/Amazon 登录态、下载过的卖家精灵导出表、生成出来的 Excel、调试 PNG 截图。
6. 保留源码、`requirements.txt`、`package.json`、lockfile、docs，以及 `chart-runtime\vendor\echarts.min.js`。
7. 新电脑上重新运行 `scripts\setup_first_run.ps1` 创建自己的 `.venv`，再做首次网页登录。

`.gitignore` 已经排除生成工作簿、PNG、JSON 输出、浏览器配置、`outputs`、`.venv` 和 `node_modules`。桌面缓存不在项目目录里，默认也不会随项目分发。

给 Git 初学者的说明：

- `.gitignore` 只控制“还没被 Git 跟踪的新文件”默认不进 Git。
- 已经被 Git 跟踪的文件，照样能修改、提交、回退。
- 被忽略的文件如果确实要提交，可以用 `git add -f 文件名`。
- 如果某个已经提交过的文件以后不想再让 Git 管，可以用 `git rm --cached 文件名`。

## 安全规则

- 绝对不要打包 `.workflow-state` 或 `.browser-state.json`，里面可能包含 Cookie 或 token。
- 不要把用户名和密码写进项目。
- 不要绕过验证码、登录验证或安全检查。
- 修改代码时只改相关模块，避免牵连无关流程。
- 除非用户明确要求，否则不要改变模板既有行高和列宽。
- 新增业务功能时，放在 `modules\功能名\` 下，并通过 `scripts\run_workflow.py` 统一调度。
