# 亚马逊选品调研自动化｜当前测试版

本版本统一执行2026-11-25确认的最新流程，供同事使用自己的商品数据测试。无需选择新旧版本。先阅读 [开始使用.txt](开始使用.txt) 和 [使用指南](ecommerce-workflow-skill/docs/使用指南.md)。

## 准备与填写

- 完整解压后，在能读写本地文件和执行命令的AI工具中打开本目录，让AI读取 [SKILL.md](ecommerce-workflow-skill/SKILL.md)。
- 先检查已有环境；可用的Python和Node依赖直接复用，缺项才安装在项目目录。推荐Windows、Python 3.12、Node.js 22/24、Chrome。
- 在项目专用浏览器安装卖家精灵插件，分别登录插件、卖家精灵网页和Amazon；脚本使用Playwright连接。
- 复制 [产品调研模板](ecommerce-workflow-skill/assets/产品调研模板.xlsx) 作工作副本。按 [模板填写说明](ecommerce-workflow-skill/docs/模板填写说明.md) 填市场调研A3的关键词反查ASIN、竞品分析A1的对比ASIN，并填入自己的商品数据源。

## 当前流程

市场调研含关键词、类目、12项指标和九张图；本地生成集合链接，无200条截断。竞品提前导出评论，等待期间截图填数；主图取网页高清原文件，类目图及主图原尺寸浮动。评价基于本批VP/Vine评论，按最新规则筛样并归纳，全部原评论合并回唯一评价数据源。

上传已填写的工作簿并要求分析后，AI会自动继续各阶段，不要求再次确认开始。站点按用户指定、商品源链接、竞品链接依次确定，无信息才默认美国；邮编自动读取已有配置。评分等空白及模块页名前缀差异不影响启动。

各阶段由AI按 [执行顺序](ecommerce-workflow-skill/docs/new_workflow_trial.md) 衔接，交付前执行 [验收规范](ecommerce-workflow-skill/docs/delivery_checks.md)。目前不是单个命令完成全部业务，不能无参数运行旧总入口。

## 测试与版本

已完成作者机器上的业务和单项验证，具体范围、已知限制及反馈方式见使用指南。12站支持指集合链接格式，当前新版完整业务主要验证美国站。新电脑需核验本机账号、插件及实际采集。

同事测试包不带Git历史、账号状态、真实商品数据、截图、旧评价缓存或依赖目录；包含已确认的空数据源模板。包内 `PACKAGE-MANIFEST.json` 标明源提交及文件SHA256。

公开GitHub的main版可能落后于此测试包，暂不要用main覆盖本目录。测试包无需GitHub账号即可使用；本地整理不代表已经推送。维护者打包方式见 [分发说明](ecommerce-workflow-skill/docs/distribution.md)。
