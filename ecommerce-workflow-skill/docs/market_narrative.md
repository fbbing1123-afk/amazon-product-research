# 市场调研分析与总结自动填写

用户提供工作簿并要求完整分析时，AI必须完成市场调研七处文字分析，不再将它们留空，也不另行请求填写许可。完成图表、竞品和评价回填后执行本阶段，再验收浮动检查点和最终全图嵌入。已有嵌入版补分析可使用同一ZIP定点入口，不整体重存文件。

## 七处来源和内容

| 结论区 | 主要证据 | 应写内容 |
| --- | --- | --- |
| 整体销售数据右侧 | 本批商品源及12指标 | 样本规模、集中度、典型销量/价格、均值与中位差异及选品参考 |
| 品牌分析右侧 | 源表品牌、月销量、月销售额，三张品牌图 | 头部品牌表现、品牌/商品集中度及竞争特点 |
| 上架时间分析右侧 | 源表上架日期、销量/销售额，三张上架时间图 | 近6年分组、新老商品表现、缺失日期及口径；截面不能推成长期增长 |
| Review数量分析右侧 | 本批卖家精灵报告评分数分布图及真实样本数 | 评分积累门槛、不同评分数区间销售表现；不能把低评分数等同新品 |
| Rating分析右侧 | 同报告评分值分布图 | 评分区间与销量分布、产品体验对标；不把相关性写成因果 |
| 价格分析右侧 | 同报告价格分布图 | 主流价格区间、销售分布、价格定位参考；不以销量推断利润 |
| 底部总结 | 前述证据，以及本页关键词、有效类目数量占比 | 市场特点、细分方向、门槛、机会与待验证风险 |

源表有效ASIN数、具有销量的统计样本、前台/插件条数、报告三图各自有效样本数分别记录。Review/Rating/价格必须使用对应报告图的口径，不能用229条源数据计算结果冒充153样本图表的结论。缺失值不补0，无证据不能编造增长、利润或必然成功等判断。

## AI自动执行步骤

以下命令由AI执行，JSON也是AI基于本批证据生成，不要求用户手工编写。

```powershell
python modules/product_research/scripts/fill_market_narrative.py prepare --workbook "全部填数检查点.xlsx" --as-of "本批采集日期YYYY-MM-DD" --report-manifest "本批三图/manifest.json" --output "本批缓存/market-evidence.json"
```

prepare输出源表指标、品牌/上架年份分组、未知值计数、关键词和类目前10占比、报告三图原路径/指纹/样本数，以及按标签动态找到的七个目标区域。prepare只是材料准备，不等于已经分析。

AI读取JSON及本批六张本地图，并实际打开三张报告原图，逐项写简洁的“数据表现＋判断＋选品参考”。图中文字不清时只写可核实的定性分布，不能猜精确数值；有同报告结构化数据时可用它复核。各节核验分子分母、时间窗口、样本数、金额单位，独立复算引用的关键比例。

各节选择关键数字、主要判断或选品建议写入highlights，片段必须在该节正文中真实存在，不能交叠、重复或整段标红。回填保留原字体字号，只将重点片段改为红色。保存本批分析JSON，其结构为：

```json
{
  "packet_digest": "本批证据中的指纹",
  "sections": {
    "overall": {"status": "analyzed", "text": "实际分析正文", "highlights": ["分析"], "evidence": ["source.overall"]},
    "brand": {"status": "analyzed", "text": "实际分析正文", "highlights": ["分析"], "evidence": ["source.brands"]},
    "listing": {"status": "analyzed", "text": "实际分析正文", "highlights": ["分析"], "evidence": ["source.listing"]},
    "review": {"status": "analyzed", "text": "实际分析正文", "highlights": ["分析"], "evidence": ["report.review"]},
    "rating": {"status": "analyzed", "text": "实际分析正文", "highlights": ["分析"], "evidence": ["report.rating"]},
    "price": {"status": "analyzed", "text": "实际分析正文", "highlights": ["分析"], "evidence": ["report.price"]},
    "summary": {"status": "analyzed", "text": "实际市场总结", "highlights": ["总结"], "evidence": ["source.overall", "source.categories", "source.keywords", "report.price"]}
  }
}
```

未取得某图或有效上架日期时，该节status必须是unavailable，正文明确说明缺什么、哪些判断无法完成，不留空、不冒充已分析。底部总结仅引用实际可用证据。可用数据不能以unavailable跳过。

```powershell
python modules/product_research/scripts/fill_market_narrative.py fill --workbook "全部填数检查点.xlsx" --packet "本批缓存/market-evidence.json" --analysis "本批缓存/market-analysis.json" --output "含市场分析检查点.xlsx"
python modules/shared/quality_gates.py --step market_narrative --workbook "含市场分析检查点.xlsx"
```

## 保存和验收

按实际标签、图框及合并区域找七个文字区域，不写死行号。不覆盖图片或公式；已有文字默认拒绝覆盖，明确重写时传--replace-existing。输入工作簿、图像或证据改变后须重新prepare，不能套其它批次或旧数据结论。

只修改市场调研页七个目标单元格的XML文本。其它ZIP部件、媒体、嵌入图片、富文本、样式、行高、列宽、合并均保持；输出独立新文件，原文件及已存在输出不能覆盖。现有字号下放不下时精简正文，不扩行扩列、不缩字号。核验七处都有实际正文或具体缺失说明，并用WPS检查显示；字面占位或只准备了JSON不能算完成。

最终门禁已包含market_narrative，完整报告缺少任一节会失败；其它实际缺图仍按原规则处理。门禁只保证结构、来源绑定和完整性，数字与判断仍由AI逐项回查，不能把门禁通过当作结论正确的唯一证明。
