---
name: AutoCaseReview
description: 自动生成"测试用例评审"相关的会议纪要与邮件文档。基于 requirement_data 目录下的需求/用例数据，按需求（ID）填充 LC-SOP-RC-007-R02 (Excel) 与 LC-SOP-RC-003-M01 (Word) 两份 SVN 归档模板（每条需求各 2 份），同时按迭代生成符合 email.png 格式的评审邮件 Word（每迭代 1 份，独立调用）。
---

# AutoCaseReview — 测试用例评审会议纪要自动生成

## 1. 用途

本 Skill 用于把 Ones/TAPD 等系统导出的"迭代需求 + 用例 + 代办事项"原始数据（`requirement_data/*.xlsx`），自动化转换为下游正式归档与沟通文档：

- **SVN 归档（按需求 ID 生成，多迭代结束后批量产出）**
  - 每一条需求（ID）各产出一份 `LC-SOP-RC-007-R02_测试用例评审_iBatchInsight_<ID>_<标题>_A1.xlsx`（含 封皮 / 签名页 / 13 项检查项评审记录）
  - 每一条需求（ID）各产出一份 `LC-SOP-RC-003-M01_会议纪要_iBatchInsight_测试用例评审_<ID>_<标题>_A1.docx`（SOP 正式会议纪要表）
  - 例：8 条需求 → 16 份 SVN 文件
- **评审邮件（按迭代生成，单独调用）**
  - 每个迭代 1 份 `<product>_<iteration>_测试用例评审_邮件.docx`（5 列精简表格 + 代办事项列表，与 `email.png` 样式一致）

## 2. 何时使用

- 用户说："生成测试用例评审会议纪要 / 用例评审邮件 / 评审纪要"。
- 用户提供了 requirement_data 目录下的 xlsx 数据文件（或一份符合格式约定的 xlsx）。
- 用户要求归档到 SVN（→ 按需求批量生成），或向团队发送评审结果邮件（→ 按迭代生成）。

## 3. 输入约定

- **模板（已内置）**：位于 `inputs/`
  - `LC-SOP-RC-007-R02_测试用例评审_iBatchInsight_需求名称_版本号.xlsx`
  - `LC-SOP-RC-003-M01_会议纪要_iBatchInsight_测试用例评审_需求名称_版本号.docx`
  - `email.png`（邮件样式参考）
- **数据源**：位于 `inputs/requirement_data/`，一份或多份 xlsx，列字段约定：
  `ID, 标题, 测试, 测试用例链接, 计划完成日期, 前端开发, 后端开发, 创建者, 所属迭代, 代办事项1@责任人, 代办事项2@责任人, 代办事项3@责任人`
  - 多份 xlsx 文件可以覆盖多个迭代，所有需求行合并后按 ID 拆分生成 SVN 文件。
  - 邮件场景必须指定一个 `所属迭代`，仅该迭代的需求行参与生成。

## 4. 输出约定

- 全部产物输出至 `outputs/`。
- 命名模板集中在 `config/filename_templates.yaml`（修改这个文件即可，不用动 Python）。默认模板：
  - SVN Excel：`LC-SOP-RC-007-R02_测试用例评审_{product}_{title}_{version}.xlsx`
  - SVN Word：`LC-SOP-RC-003-M01_会议纪要_{product}_测试用例评审_{title}_{version}.docx`
  - 邮件 Word：`{product}_{iteration}_测试用例评审_邮件.docx`
- 可用占位符：`{product}` / `{title}` / `{iteration}` / `{version}` / `{req_id}`。所有占位符值会自动把 `/ \ : * ? " < > |` 替换为 `_`，标题截断到 60 字符。未在上下文中提供的占位符会被替换为空字符串。
- **封面 D3 单元格 "版本：A1" 是 SOP 表单自身的修订号，保持模板默认不动**；文件名末尾的 `{version}` 才是项目发布版本号（如 `v1.0.1` / `SP7`），由用户在生成前指定。

## 5. 字段映射（核心）

每行末尾标注配置位置（**改 yaml 即时生效，不用动 py**）：

| 模板位置 | 数据来源 / 默认值 | 配置位置 |
|---|---|---|
| Excel · 封皮 / 签名页（C2/D2/D3 等） | **完全保持模板原样，不动** | — |
| Excel · `产品名称` | `iBatchInsight` | `--product` CLI |
| Excel · `Ones需求` | 当前需求行 `ID + 标题` | — |
| Excel · `测试用例链接` | 当前需求行的 `测试用例链接` 列 | `content_rules.yaml::excel.case_link` |
| Excel · `发起人` | 当前需求行的 `测试` 列 | `content_rules.yaml::excel.initiator` |
| Excel · `主持人` | 当前需求行的 `测试` 列 | `content_rules.yaml::excel.host` |
| Excel · `评审人` | `测试 + 创建者 + 前端开发 + 后端开发` 去重 | `content_rules.yaml::excel.reviewer` |
| Excel · `评审时间` | 计划完成日期**前一天** | `content_rules.yaml::excel.review_date` |
| Excel · 13 项检查项 `检查结果` | 全部 `通过` | py 硬编码（路线图） |
| Word(SVN) · `会议名称` | `{title} 测试用例评审会议纪要` | `content_rules.yaml::meeting.name_template` |
| Word(SVN) · `会议地点` | `线上会议` | `content_rules.yaml::meeting.place` |
| Word(SVN) · `会议时间` | 计划完成日期**前一天** | `content_rules.yaml::meeting.time` |
| Word(SVN) · `记录人员` | CLI 显式指定 | `--recorder` |
| Word(SVN) · `参会人员` | 当前需求的 `测试 + 前端开发 + 后端开发 + 创建者` 去重 | `content_rules.yaml::participants.source_columns` / `separator` |
| Word(SVN) · `会议内容` | 自动汇总 + 逐条列出代办事项（移除 `@责任人`） | `content_rules.yaml::meeting.content.list_todos` / `strip_mentions` |
| Word(SVN) · todo `问题描述` | `代办事项N@责任人` 原文 | — |
| Word(SVN) · todo `责任人` | 从代办事项文本中提取的 `@姓名` | py 正则 |
| Word(SVN) · todo `计划解决日期` | 计划完成日期**前一天** | `content_rules.yaml::todo.plan_date` |
| Word(SVN) · todo `状态` | `已完成` | `content_rules.yaml::todo.status` |
| Word(SVN) · todo `备注` | 空 | `content_rules.yaml::todo.note` |
| Word(邮件) · 5 列表 | `ID / 标题 / 测试 / 通过 / 空` | py 硬编码 |
| Word(邮件) · 代办列表 | 按需求分组 bullet，保留 `@责任人` 原文 | py 硬编码 |

**覆盖优先级**：CLI 显式参数 > `content_rules.yaml` > py 默认值。

**配置文件清单**（`config/`）：
- `filename_templates.yaml` — 文件名模板 + 输出子目录映射（见 §4）
- `content_rules.yaml` — 文档内容规则（见上表，含 `meeting` / `participants` / `todo` / `excel` 四节）

**输出目录**（`outputs/`）按文件类型分目录：
- `outputs/测试用例评审/` — 所有 SVN Excel
- `outputs/会议纪要/` — 所有 SVN Word 会议纪要
- `outputs/邮件/` — 所有评审邮件 Word

## 6. 执行流程

0. **生成前必问**：调用 `AskUserQuestion` 询问"当前版本号是什么？"（例如 `v1.0.1` / `SP7` / `CPV-SP7-v1.2`）。该版本号会作为 SVN 文件名末尾的 tag，**禁止默认 `A1`**（A1 是 SOP 表单修订号，已写在封面 D3，不可混用）。若用户只说了 `--mode email`，则跳过此询问。
1. 扫描 `inputs/requirement_data/` 下的 xlsx，合并为统一 DataFrame。
2. 按模式分支：
   - `--mode svn`：遍历每一条需求行 → 各产出 1 份 Excel + 1 份 Word；会议元信息（主持人/评审人/参会人员等）跨需求共享。
   - `--mode email --iteration <X>`：筛选该迭代所有需求 → 产出 1 份邮件 Word（不需要版本号）。
   - `--mode all`：先按需求批量产出 SVN 文件（带版本号），再按 `--iteration` 产出邮件 Word。
3. 读取模板，按字段映射填充，**保留模板原有样式、合并单元格、签名页与封皮（包括 D3=版本：A1）**。
4. 输出到 `outputs/`，返回生成的文件清单。
5. 若有关键缺失字段（如评审时间、主持人），用 `AskUserQuestion` 询问后再生成。

## 7. 注意事项

- 模板的封皮 / 签名页 / 13 项检查项结构 **不得删除或重排**，只填充数据。
- 邮件 Word 中的 `@责任人` 必须保留原文，不得拆分。
- SVN 文件按需求生成，**不要把多条需求合并到同一个 SVN 文件**（每条需求对应独立的 SOP 归档记录）。
- 邮件 Word 必须按迭代生成，**不能跨迭代**。
- 任何无法自动判定的字段（如评审结论是否"通过"、参会人员名单）走用户问答。

## 8. 依赖

`openpyxl`、`python-docx`、`pandas`、`pyyaml`（首次运行前自动安装）。
