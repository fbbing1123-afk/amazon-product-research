# 亚马逊选品调研自动化

基于 Windows、Python、Node.js 和浏览器自动化，将商品数据加工为同一份 Excel 调研报告。流程涵盖关键词、类目、市场、评论及竞品分析。

## 当前状态

- 源码基线来自 `亚马逊选品自动化-开发交接-20261121.zip`，于 2026-11-22 建立本地 Git 管理。
- 当前默认“最新周 PPC 竞价＋关键词按旺季搜索量降序”，原旺季区间算法可通过参数选择。两种方案的区别、开发记录及恢复方式见[开发日志](开发日志.md)。
- 类目月总销量下方显示产品占比百分比，并保留蓝色/粉色底色；支持从独立数据源工作簿局部刷新已有报告，详见[类目规则](ecommerce-workflow-skill/docs/market_research_category_section_rules.md)。
- 关键词竞价刷新已用8个关键词实测，相关26项测试通过；其它业务模块和整条工作流尚未在本机重新验收。
- 类目截图边缘裁切及部分结果问题仍需实测；保留固定等待满 20 秒策略。
- GitHub 私有仓库：[amazon-product-research](https://github.com/fbbing1123-afk/amazon-product-research)，源码和版本历史用于跨电脑同步及授权同事下载。
- 仓库不包含业务 Excel、原始模板、历史输出、登录状态或依赖目录。

## 文件入口

| 路径 | 用途 |
| --- | --- |
| [AGENTS.md](AGENTS.md) | 开发助手的协作与 GitHub 同步约定 |
| [开发交接.md](开发交接.md) | 2026-11-21 交接快照、迁移背景和待验证事项 |
| [codex.md](codex.md) | 工作簿格式与交付验收规范 |
| [使用指南](ecommerce-workflow-skill/docs/使用指南.md) | 环境准备及日常使用 |
| [SKILL.md](ecommerce-workflow-skill/SKILL.md) | 业务流程和模块运行规则 |
| [CHANGELOG.md](CHANGELOG.md) | 已记录的项目变更 |
| [开发日志.md](开发日志.md) | 两种竞价方案、验证记录、备份标记与恢复方式 |

`开发交接.md` 中“尚未建立 Git 仓库”的描述属于历史状态；当前状态以本 README 和实际 Git 记录为准。

## 新电脑安装

推荐 Windows、Python 3.12、Node.js 22/24 和 Chrome。原生切片器等步骤需要本机 Excel/WPS COM 支持。

克隆或下载仓库后，进入 `ecommerce-workflow-skill` 目录，再运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup_first_run.ps1
powershell -ExecutionPolicy Bypass -File scripts\check_environment.ps1 -RequireInstalled
powershell -ExecutionPolicy Bypass -File scripts\open_sellersprite_login_browser.ps1 -Site US
```

在打开的项目浏览器中安装 softtime/sorftime 插件，登录卖家精灵和目标 Amazon 站点，确认配送邮编。依赖、插件和登录状态需要每台电脑单独准备，不通过 Git 同步。

环境自检不替代账号、插件、COM 及真实工作簿验证。不要上传账号密码、Cookie 或验证码。

## 工作流程

```text
输入工作簿 → 关键词调研 → 类目调研 → 市场分析 → 评论调研
          → AI 竞品语义分析 → Excel 回填、差评透视和切片器 → 最终验收
```

`scripts/run_workflow.py` 串联前四个脚本阶段。竞品语义分析需要 AI 完成交接 JSON，再进入竞品后处理；脚本阶段结束不代表完整报告验收通过。实际使用需另行提供业务工作簿、站点、邮编及运行范围。

## 开发和同步

1. 开始工作前检查分支和未提交修改；有远程仓库后先获取最新状态。
2. 使用功能分支进行新功能开发，完成相应检查和实际验证。
3. 本地按功能提交，说明具体变化。
4. 功能完成并跑通后，向用户报告并询问是否同步 GitHub；确认后才推送。
5. 更换电脑前确认需要同步的提交已推送；另一台电脑开始修改前拉取。

业务数据、检查点和输出文件默认被忽略，更换电脑继续同一个案例时需要另行安全转移。私有仓库的同事需要先获得访问权限；经验证的版本可通过 Release 提供下载。
