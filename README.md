# 亚马逊选品调研自动化｜当前测试版

本版本为20261127-test1，执行截至2026-11-27确认的最新流程；包含类目前10和市场分析自动填写，旧包不会自动更新。无需选择新旧流程。先阅读 [开始使用.txt](开始使用.txt) 和 [使用指南](ecommerce-workflow-skill/docs/使用指南.md)。

## 准备与填写

- 完整解压后，在能读写本地文件和执行命令的AI工具中打开本目录，让AI读取 [SKILL.md](ecommerce-workflow-skill/SKILL.md)。
- 先检查已有环境；可用的Python和Node依赖直接复用，缺项才安装在项目目录。推荐Windows、Python 3.12、Node.js 22/24、Chrome。
- 项目浏览器自动选择空闲端口并保存，各阶段直接读取实际地址。在该浏览器安装卖家精灵插件，分别登录插件、卖家精灵网页和Amazon；脚本使用Playwright连接。
- 复制 [产品调研模板](ecommerce-workflow-skill/assets/产品调研模板.xlsx) 作工作副本。按 [模板填写说明](ecommerce-workflow-skill/docs/模板填写说明.md) 填市场调研A3的关键词反查ASIN、竞品分析A1的对比ASIN，并填入自己的商品数据源。

## 当前流程

市场调研含关键词、数量占比前10个类目、12项指标和九张图；AI自动填写七处市场分析及总结，关键数字、判断和建议局部标红，见[分析规则](ecommerce-workflow-skill/docs/market_narrative.md)；本地生成集合链接，无200条截断。竞品默认最多2个同时开页及导出评论，等待期间截图填数，完成后补下一个；主图取网页高清原文件，类目图、主图及三张市场报告截图先原尺寸浮动定位。全部填数结束后[统一嵌入图片](ecommerce-workflow-skill/docs/final_image_embedding.md)，不调整行高、列宽或合并区域。评价基于本批VP/Vine评论，全部原评论合并回唯一评价数据源。

上传已填写的工作簿并要求分析后，AI会自动继续各阶段，不要求再次确认开始。站点按用户指定、商品源链接、竞品链接依次确定，无信息才默认美国；邮编自动读取已有配置。评分等空白及模块页名前缀差异不影响启动。

各阶段由AI按 [执行顺序](ecommerce-workflow-skill/docs/new_workflow_trial.md) 衔接，交付前执行 [验收规范](ecommerce-workflow-skill/docs/delivery_checks.md)。目前不是单个命令完成全部业务，不能无参数运行旧总入口。

Amazon已知确认页按[接入说明](ecommerce-workflow-skill/docs/amazon_checkpoint.md)在业务内唤醒目标页、尝试勾选并继续。确认通过却跳到其它普通页面时自动重进原目标一次，随后核验目标、邮编及插件；仍失败则保留进度。已启动评论导出不重复触发。已使用用户Excel进行美国站真实试跑，实际确认恢复及剩余限制见接入说明；本轮代码纳入本地Git测试版，未推送GitHub。

## 测试与版本

本次改动、验证范围、已知限制和同事测试步骤见[测试版说明](ecommerce-workflow-skill/docs/test_release_20261127.md)，日常操作见使用指南。最终图片采用WPS单元格图片格式，请用WPS查看。12站支持指集合链接格式，当前新版完整业务主要验证美国站。新电脑需核验本机账号、插件及实际采集。

同事测试包不带Git历史、账号状态、真实商品数据、截图、旧评价缓存或依赖目录；包含已确认的空数据源模板。包内 `PACKAGE-MANIFEST.json` 标明源提交及文件SHA256。

公开GitHub的main版可能落后于此测试包，暂不要用main覆盖本目录。测试包无需GitHub账号即可使用；本地整理不代表已经推送。维护者打包方式见 [分发说明](ecommerce-workflow-skill/docs/distribution.md)。
