# AutoCaseReview

> 自动生成 **测试用例评审** 的 SVN 归档文件与评审邮件。SVN 归档按**需求（ID）**粒度生成（每条需求各 1 份 Excel + 1 份 Word），评审邮件按**迭代**粒度生成。

## 功能概览

| # | 产物 | 粒度 | 用途 |
|---|---|---|---|
| 1 | `LC-SOP-RC-007-R02_测试用例评审_<ID>_<标题>_*.xlsx` | **每条需求 1 份** | SVN 归档（13 项检查项评审记录） |
| 2 | `LC-SOP-RC-003-M01_会议纪要_*_<ID>_<标题>_*.docx` | **每条需求 1 份** | SVN 归档（SOP 会议纪要表） |
| 3 | `<product>_<iteration>_测试用例评审_邮件.docx` | **每迭代 1 份** | 群发评审邮件（独立调用） |

> 例：1 个迭代包含 8 条需求 → SVN 归档产出 16 份（8 × 2），邮件产出 1 份。

## 目录结构

```
AutoCaseReview/
├── main.py                          # 项目根入口（占位）
├── .claude/skills/AutoCaseReview/
│   ├── SKILL.md                     # Skill 定义（含字段映射）
│   ├── README.md                    # 本文件
│   ├── inputs/                      # 模板与原始数据
│   │   ├── LC-SOP-RC-007-R02_*.xlsx        # Excel 模板
│   │   ├── LC-SOP-RC-003-M01_*.docx        # Word 会议纪要模板
│   │   ├── email.png                       # 邮件样式参考
│   │   └── requirement_data/               # 真实数据 xlsx（可含多个迭代）
│   ├── outputs/                     # 生成结果（按类型分子目录）
│   │   ├── 测试用例评审/                # SVN Excel (LC-SOP-RC-007-R02)
│   │   ├── 会议纪要/                    # SVN Word  (LC-SOP-RC-003-M01)
│   │   └── 邮件/                       # 评审邮件 Word
│   ├── scripts/                     # 生成脚本
│   │   ├── data_loader.py           # 数据加载 / 合并 / 按需求迭代
│   │   ├── generate_svn_excel.py    # LC-SOP-RC-007-R02（按需求）
│   │   ├── generate_svn_word.py     # LC-SOP-RC-003-M01（按需求）
│   │   ├── generate_email_word.py   # 评审邮件（按迭代）
│   │   └── main.py                  # CLI 统一入口
│   └── config/
│       ├── filename_templates.yaml  # 文件名模板 + 输出子目录映射（可编辑）
│       └── content_rules.yaml       # 文档内容规则（可编辑）
```

## 数据源格式（`inputs/requirement_data/*.xlsx`）

| 字段 | 说明 |
|---|---|
| `ID` | Ones 用例编号，如 `#203726`（**SVN 文件按此拆分**） |
| `标题` | 需求/用例标题，如 `【数据分析工作台】能力分析等级配置` |
| `测试` | 测试负责人 |
| `测试用例链接` | Ones URL |
| `计划完成日期` | YYYY-MM-DD（自动剥离 `00:00:00`） |
| `前端开发` / `后端开发` / `创建者` | 参会人员来源（跨需求汇总去重） |
| `所属迭代` | 如 `CPV-SP7`，**邮件文件按此拆分** |
| `代办事项1@责任人` / `2` / `3` | 评审记录中需要落实的待办，包含 `@责任人` 原文 |

## 使用方式

> **生成 SVN 文件前，必须先问用户当前版本号**（如 `v1.0.1` / `SP7`）。Skill 在执行前会用 `AskUserQuestion` 询问，并把它作为 `--version` 传给脚本。

```bash
cd .claude/skills/AutoCaseReview/scripts

# 1) SVN 归档：按需求批量生成（8 条需求 → 16 份文件）。--version 必填。
python3 main.py --mode svn \
  --version "v1.0.1" \
  --meeting-time "2026-06-25 14:00" --recorder "黄美玲" \
  --initiator "荆慧慧" --host "荆慧慧" \
  --reviewer "张三" --review-date "2026-06-25"

# 2) 邮件：按迭代生成（必须指定 --iteration；不需要版本号）
python3 main.py --mode email --iteration CPV-SP7

# 3) 两者一起（SVN 带版本号，邮件不需要）
python3 main.py --mode all --iteration CPV-SP7 \
  --version "v1.0.1" \
  --meeting-time "2026-06-25 14:00" --recorder "黄美玲" \
  --initiator "荆慧慧" --host "荆慧慧" \
  --reviewer "张三" --review-date "2026-06-25"
```

## 文件命名规则

**所有命名模板集中在 `config/filename_templates.yaml` —— 改这个文件即可，不用动 Python 代码。**

默认模板：

```yaml
svn_excel:  "LC-SOP-RC-007-R02_测试用例评审_{product}_{title}_{version}.xlsx"
svn_word:   "LC-SOP-RC-003-M01_会议纪要_{product}_测试用例评审_{title}_{version}.docx"
email_word: "{product}_{iteration}_测试用例评审_邮件.docx"
```

可用占位符：

| 占位符 | 含义 | 示例 |
|---|---|---|
| `{product}` | 产品名 | `iBatchInsight` |
| `{title}` | 当前需求标题（已安全化、截断 60 字符） | `【数据分析工作台】能力分析等级配置` |
| `{iteration}` | 所属迭代；跨多迭代时为 `多迭代` | `CPV-SP7` |
| `{version}` | 用户在生成前指定的版本号（仅 SVN） | `v1.0.1` |
| `{req_id}` | 当前需求 ID（去掉 `#`） | `203726` |

> - 文件名**不含 ID**（默认模板）—— 如需加回，把 `{req_id}` 写进 yaml。
> - 所有占位符值会自动把 `/ \ : * ? " < > |` 替换为 `_`。
> - `<version>` 由用户在生成前指定，脚本不提供默认值。
> - 封面 D3 单元格的 "版本：A1" 是 SOP 表单自身的修订号，与文件名末尾的版本号**无关**，保持模板默认不动。

## 字段映射要点

详见 `SKILL.md` 第 5 节。摘要：
- Excel 模板 3 个 Sheet（封皮 / 签名页 / 测试用例评审记录 13 项检查项）保留不动，只填数据。
- Word(SVN) 的"改进或遗留工作项"由当前需求的 `代办事项N@责任人` 拆行展开。
- Word(邮件) 5 列表 = `ID / 标题 / 测试 / 评审结果(=通过) / 备注#(空)`，覆盖该迭代全部需求。
- 参会人员、主持人、评审人等会议元信息**跨所有需求共享**（同一场评审会）。

## 依赖

```bash
pip3 install python-docx openpyxl pandas pyyaml
```

## 路线图

- [x] `config/filename_templates.yaml` 文件名模板外置
- [ ] `config/field_mapping.yaml` 字段映射外置
- [ ] 13 项检查项 "通过/不通过" 自动判定（基于 `@修改用例` 标记）
- [ ] 跨多个迭代 xlsx 的批量归档验证（待用户提供多样本数据）

## 约束

- 严禁修改模板的封皮 / 签名页 / 13 项检查项结构。
- `@责任人` 原文必须保留。
- **SVN 文件按需求生成**，不得把多条需求合并到同一份 SVN 文件。
- **邮件 Word 按迭代生成**，不得跨迭代。
