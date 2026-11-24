# 卖家精灵评论评价分析新版

这是卖家精灵“评论下载”数据源对应的独立分析入口。用户于 2026-11-24 确认以下规则。旧网页逐条抓取版保留，但新版不得调用它的抽样、痛点词库、语义缓存、覆盖率或差评透视流程。

## 筛选口径

每个下载目标 ASIN 单独处理，始终保留原 Excel。

1. 保留 `VP评论=Y` **或者** `Vine Voice评论=Y` 的评论，同时满足时只计一条。
2. 按赞同数数值降序；赞同数相同按源表行号升序，结果可重复。
3. 优点：筛选 4、5 星，取前 **30** 条；不足时全部使用。
4. 缺点：筛选 1、2、3 星，取前 **100** 条；不足时全部使用。
5. 缺点交给全新、独立上下文的子智能体，核心要求原样使用：**“基于我给的评论，分点精简总结产品的差评原因，每点精简在 15 个字以内。”**

赞同数空白视为 0；缺少整列或出现负数、缩写、非整数时停止并核对，不能把未知值冒充 0。空白 ASIN 按之前约定保留，仍属于本次下载目标的导出材料；其它子体 ASIN 也保留原值并提示，不能由语义词库擅自重新分配商品。原始记录不可用其它商品的评论补齐。

先选样再分析。标题和正文都为空的记录保留编号，不能引用为结论证据，也不从低赞评论补位。离题、无明确原因的内容不编造结论。优点沿用每点最多 20 字；缺点每点最多 15 字，均为每点一行、不编号、不附旧版覆盖率。不人为限制结论条数，近义内容合并；证据序号只保存在 JSON，不显示到优缺点单元格。

## 准备编号评论

在 skill 根目录用已配置的 Python 执行：

```powershell
python modules/competitor_analysis/sellersprite_review_analysis.py prepare --asin B0XXXXXXXX --export "评论原始导出.xlsx" --output-dir "outputs/review-analysis"
```

每次生成独立目录，包括：

- `packet.json`：只读原始来源、SHA-256、筛选规则、数量、完整选样和编号。
- `advantages.prompt.txt`：当前商品的好评材料。
- `disadvantages.prompt.txt`：当前商品的差评材料及子智能体要求。

没有读取旧痛点 JSON、旧差评页、翻译缓存或其它商品分析结果的路径。插件导出的译文可供参考，语义冲突时按同条评论的原文判断。

## 主智能体编排

1. 主智能体阅读本次 `advantages.prompt.txt`，按其格式保存 `advantages.result.json`。
2. 为当前商品创建专门的差评子智能体，**不继承主对话历史**（支持时使用 `fork_turns="none"`）。把本次 `disadvantages.prompt.txt` 的全文作为消息发给它。不得传旧角色 `agent_prompt.md`、旧痛点分组、历史总结或多个商品混合的评论。
3. 子智能体只读这批按序号排列的原文和译文，不联网、不搜索、不读项目文件、不改代码，不创建其它子智能体；返回 JSON。主智能体原样保存到 `disadvantages.result.json`。
4. 逐点核对结论与所引评论是否吻合，近义点是否重复、是否把“担心会损坏”写成“已经损坏”。有问题时让该子智能体只按本次证据修正，不用旧词库纠正它。
5. 每个商品使用独立子智能体上下文和结果文件。Python 只负责准备、校验、写表，不能把生成提示文件说成已经实际调用了子智能体。环境没有子智能体能力时，应明确说明尚未完成缺点分析，不静默走回旧版。

结果需包含本次 `asin`、`packet_digest`、`kind`、`reviewed_numbers` 和 `points`。每条 point 包含 `text` 及 `review_numbers`。来源或样本发生变化后必须重新分析，不能改哈希来复用旧结果。校验能证明序号、长度和批次匹配，不能自动证明语义正确，主智能体必须核对原文。

没有符合条件的评论时，主智能体可保存空列表与空 `reviewed_numbers`，不必派空任务。存在样本却无明确可支持结论时，结果必须填写 `no_supported_points_reason`；写表为 `/`，具体原因保留在校验结果 JSON，不能断言商品没有缺点。

## 校验和回填

```powershell
python modules/competitor_analysis/sellersprite_review_analysis.py check --packet "本次目录/packet.json" --advantages "本次目录/advantages.result.json" --disadvantages "本次目录/disadvantages.result.json" --output "本次目录/checked-analysis.json"

python modules/competitor_analysis/sellersprite_review_analysis.py fill --packet "本次目录/packet.json" --advantages "本次目录/advantages.result.json" --disadvantages "本次目录/disadvantages.result.json" --workbook "竞品工作副本.xlsx" --output "新版评价工作副本.xlsx"
```

`check` 与 `fill` 均重新读取原始导出，核对样本、排序和来源；发现错商品、旧批次、漏读序号、越界证据、重复结论或超长文字即失败。`fill` 每次处理一个商品，多商品按顺序在工作副本上继续回填，不能相互覆盖输出。

写表按“详情页链接”的 ASIN 找列，按“优点”“缺点”实际标签找行，支持前方模块增行后的定位。仅修改该商品两个单元格的文本，保留样式、行高、列宽及其它 XML 部件；拒绝覆盖原表、已存在的输出或猜测缺失/合并单元格。使用 ZIP/XML 局部更新并验收，避免重存工作簿破坏 WPS 图片或原生切片器。

本入口只负责本次约定的优缺点，不生成品类总结、改进建议或差评透视，也未改总调度默认版本。不得随后运行旧 `competitor_analysis/run.py` 将旧缓存结果覆盖回来。整份业务报告交付仍需相应模块门禁和视觉验收；这里的两格回填测试不能代替完整工作流验收。

源码、测试和规则说明纳入 Git；真实评论表、编号材料、智能体结果和回填工作簿仅保存在本地输出目录，不提交到公开仓库。
