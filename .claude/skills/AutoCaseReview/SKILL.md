---
name: AutoCaseReview
description: 自动生成"测试用例评审"相关的会议纪要与邮件文档。基于 requirement_data 目录下的需求/用例数据，按需求（ID）填充 LC-SOP-RC-007-R02 (Excel) 与 LC-SOP-RC-003-M01 (Word) 两份 SVN 归档模板（每条需求各 2 份），同时按迭代生成符合 email.png 格式的评审邮件 Word（每迭代 1 份，独立调用）。
---

# AutoCaseReview Skill

用于从 `inputs/requirement_data/*.xlsx` 生成测试用例评审邮件和 SVN 归档会议纪要。

## 触发规则
- 用户说 `生成SPx邮件用例评审会议纪要`：从提示词中提取 SP 关键字，如 `SP7`、`SP8`、`SP9`。只读取文件名包含该关键字的 `requirement_data` xlsx，生成 1 份邮件 Word，不询问版本号。
- 用户说 `生成SVN用例评审会议纪要`：读取 `requirement_data` 下所有 xlsx，按每条需求生成 SVN Excel + SVN Word。执行前必须询问当前版本号，并作为 `--version` 传入。
- 用户泛泛要求“测试用例评审会议纪要 / 用例评审邮件 / 评审纪要”时，先判断是邮件还是 SVN；无法判断时向用户确认。

## 命令
在 `.claude/skills/AutoCaseReview/scripts` 下执行：

```bash
# 邮件：<SP_KEYWORD> 来自用户提示词。
python main.py --mode email --data-file-keyword <SP_KEYWORD>

# SVN：<VERSION> 必须由用户确认，不能默认 A1。
python main.py --mode svn --version "<VERSION>"

# 只校验数据，不生成文件。
python main.py --mode email --data-file-keyword <SP_KEYWORD> --validate-only
```

如果需要使用项目根目录的虚拟环境，可执行：

```bash
.\.venv\Scripts\python.exe .claude\skills\AutoCaseReview\scripts\main.py --mode email --data-file-keyword SP8
```

## 数据选择
- 模板文件位于 `inputs/templates/`，包括 SVN Excel 模板、SVN Word 模板和 `email.png` 样式参考。
- 邮件模式：按文件名关键字筛选 `inputs/requirement_data/*.xlsx`，再从筛选后的数据推断唯一 `所属迭代`；提示词里的 `SPx` 只用于筛文件，邮件正文和文件名里的迭代只取数据列 `所属迭代`。
- SVN 模式：不按文件名筛选，读取 `inputs/requirement_data/` 下所有非 `~$` 开头的 xlsx。
- `代办事项N@责任人` 是动态列，支持任意数量，按 N 从小到大展开。
- 生成前会自动校验数据；校验失败时不生成文件，警告信息需要在回复中提示用户。

## 输出
- `outputs/邮件/`：每个迭代 1 份评审邮件 Word。
- `outputs/测试用例评审/`：每条需求 1 份 `LC-SOP-RC-007-R02` Excel。
- `outputs/会议纪要/`：每条需求 1 份 `LC-SOP-RC-003-M01` Word。
- 同名文件直接覆盖；如果文件被 Word/WPS/Excel 占用，提示用户关闭后重试。
- 触发邮件生成前，先清空 `outputs/邮件/`。
- 触发 SVN 生成前，先清空 `outputs/测试用例评审/` 和 `outputs/会议纪要/`。
- `--validate-only` 只校验数据，不清空输出目录。

## 字段规则
| 产物 | 字段 | 规则 |
|---|---|---|
| SVN Excel | 封皮 / 签名页 / 13 项检查项结构 | 保持模板原样，不删除、不重排 |
| SVN Excel | `测试用例链接` | 当前需求行 `测试用例链接` |
| SVN Excel | `发起人` / `主持人` | 默认取当前需求行 `测试` |
| SVN Excel | `评审人` | `测试 + 创建者 + 前端开发 + 后端开发` 去重 |
| SVN Excel | `评审时间` | `计划完成日期` 前一天 |
| SVN Word | `会议名称` | `{title} 测试用例评审会议纪要` |
| SVN Word | `会议地点` | 默认 `线上会议` |
| SVN Word | `会议时间` | `计划完成日期` 前一天 |
| SVN Word | `记录人员` | 默认取当前需求行 `测试`；CLI `--recorder` 可覆盖 |
| SVN Word | `参会人员` | `测试 + 前端开发 + 后端开发 + 创建者` 去重 |
| SVN Word | `会议内容` | 当前需求所有 `代办事项N@责任人` 原文；无代办填 `无` |
| SVN Word | `改进或遗留工作项` | 每条代办 1 行；责任人从 `@姓名` 提取；无代办填 `无` |
| 邮件 Word | 需求表 | `ID / 标题 / 测试 / 通过 / 空备注` |
| 邮件 Word | 章节标题 | `1.`、`2.`、`3.` 加粗；`3.评审记录/代办事项` 下的需求 ID 和标题不加粗且字号更小 |
| 邮件 Word | 代办列表 | 按需求分组，保留 `@责任人` 原文 |

配置优先级：CLI 显式参数 > YAML 配置。默认值集中维护在 `config/project.yaml`、`config/templates.yaml`、`config/filename_templates.yaml`、`config/content_rules.yaml`；不要在 Python 中另写一份默认业务配置。

## 必须遵守
- SVN 文件必须按需求生成，不得把多条需求合并到同一份 SVN 文件。
- 邮件 Word 必须按单个迭代生成，不得跨迭代。
- SVN 版本号必须来自用户确认；不要把模板封面 D3 的 `版本：A1` 当成输出文件版本号。用户输入纯数字版本（如 `1.0.2`）时，脚本会自动补成 `v1.0.2`。
- Excel 模板封皮、签名页、图片、嵌入对象和 13 项检查项必须保留。
- 邮件和会议内容中的 `@责任人` 原文必须保留。
