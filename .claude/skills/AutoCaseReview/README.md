# AutoCaseReview

自动生成 **测试用例评审** 相关文档：SVN 归档文件按需求生成，评审邮件按迭代生成。

## 目录结构
项目主体是一个 Cursor/Claude Skill，所有模板、数据、脚本和输出都集中在 `.claude/skills/AutoCaseReview/` 下。日常使用时主要关注 `inputs/requirement_data/`、`outputs/` 和 `scripts/main.py`。

```text
.claude/skills/AutoCaseReview/
├── SKILL.md                         # Agent 触发规则与执行约束
├── README.md                        # 项目说明文档
├── inputs/                          # 输入区：模板与原始需求数据
│   ├── templates/                   # 固定模板，不参与数据筛选
│   │   ├── LC-SOP-RC-007-R02_*.xlsx # SVN Excel 模板
│   │   ├── LC-SOP-RC-003-M01_*.docx # SVN Word 会议纪要模板
│   │   └── email.png                # 邮件样式参考图
│   └── requirement_data/            # 需求/用例原始 xlsx，生成时读取这里
├── outputs/                         # 输出区：同名文件会被覆盖
│   ├── 测试用例评审/                  # 生成的 SVN Excel
│   ├── 会议纪要/                      # 生成的 SVN Word
│   └── 邮件/                         # 生成的评审邮件 Word
├── scripts/                         # 生成脚本
│   ├── data_loader.py               # 读取、清洗、筛选 requirement_data
│   ├── generate_svn_excel.py        # 生成 LC-SOP-RC-007-R02 Excel
│   ├── generate_svn_word.py         # 生成 LC-SOP-RC-003-M01 Word
│   ├── generate_email_word.py       # 生成评审邮件 Word
│   └── main.py                      # CLI 统一入口
└── config/                          # 可配置项
    ├── project.yaml                 # 产品名、提示词关键字规则等项目级默认值
    ├── templates.yaml               # 输入模板文件名
    ├── filename_templates.yaml      # 输出文件名和子目录模板
    └── content_rules.yaml           # 字段来源、会议内容、todo 默认规则
```

## 工作流程
```text
① 从 Ones 导出需求数据
    - 在 Ones 中按迭代筛选需求
    - 导出字段：ID、标题、测试、测试用例链接、计划完成日期、前端开发、后端开发、创建者、所属迭代
    |
    ▼
② 手动补充代办事项
    - 在导出的 xlsx 中新增并填写：代办事项1@责任人、代办事项2@责任人...
    - 例如：UCL、LCL受小数保留位数影响 @jhh 补充prd  @hml 修改用例
    - 没有待办则留空即可
    |
    ▼
③ 将整理后的 xlsx 放到 inputs/requirement_data/
    - 文件名必须带迭代关键字，例如 SP7、SP8
    |
    ▼
用户输入提示词
    |
    ▼
④ AI 判断生成类型
    ├─ 生成SPx邮件用例评审会议纪要
    │     |
    │     ▼
    │  ⑤ AI 提取 SP 关键字
    │     例如：SP7 / SP8 / SP9
    │     |
    │     ▼
    │  ⑥ 只读取 requirement_data 下文件名包含该关键字的 xlsx
    │     默认命令：main.py --mode email --data-file-keyword <SP关键字>
    │     |
    │     ▼
    │  ⑦ 从数据中推断唯一 所属迭代
    │     |
    │     ▼
    │  ⑧ 生成评审邮件 Word
    │     输出到 outputs/邮件/
    │
    └─ 生成SVN用例评审会议纪要
          |
          ▼
       ⑤ AI 询问当前版本号
          例如：v1.0.1
          |
          ▼
       ⑥ 读取 requirement_data 下所有 xlsx
          默认命令：main.py --mode svn --version <版本号>
          |
          ▼
       ⑦ 按每条需求 ID 拆分
          |
          ▼
       ⑧ 每条需求生成 2 份 SVN 归档文件
          ├─ Excel：outputs/测试用例评审/
          └─ Word ：outputs/会议纪要/
```

## 触发方式
| 提示词 | 数据源 | 产物 | 是否需要版本号 |
|---|---|---|---|
| `生成SPx邮件用例评审会议纪要` | 只读取 `inputs/requirement_data/` 下文件名包含提示词中 SP 关键字的 xlsx | 先清空 `outputs/邮件/`，再生成 1 份邮件 Word | 否 |
| `生成SVN用例评审会议纪要` | 读取 `inputs/requirement_data/` 下所有 xlsx | 先清空 `outputs/测试用例评审/` 和 `outputs/会议纪要/`，再按每条需求生成 Excel + Word | 是 |

## 命令示例
```bash
cd .claude/skills/AutoCaseReview/scripts

# 邮件：从提示词提取 SP 关键字后传入，例如 SP8。
python main.py --mode email --data-file-keyword SP8

# 邮件：也可直接按所属迭代筛选。
python main.py --mode email --iteration CPV-SP8

# SVN：读取所有 requirement_data，--version 必填。
python main.py --mode svn --version "v1.0.1"

# 只校验数据，不生成文件。
python main.py --mode email --data-file-keyword SP8 --validate-only

# 如需保留旧输出，可显式跳过生成前清理。
python main.py --mode email --data-file-keyword SP8 --no-clear-output
```

## 产物
| 产物 | 粒度 | 输出目录 | 用途 |
|---|---|---|---|
| `LC-SOP-RC-007-R02_测试用例评审_{product}_{title}_{version}.xlsx` | 每条需求 1 份 | `outputs/测试用例评审/` | SVN 归档，测试用例评审记录 |
| `LC-SOP-RC-003-M01_会议纪要_{product}_测试用例评审_{title}_{version}.docx` | 每条需求 1 份 | `outputs/会议纪要/` | SVN 归档，SOP 会议纪要 |
| `{product}_{iteration}_测试用例评审_邮件.docx` | 每个迭代 1 份 | `outputs/邮件/` | 评审结果邮件 |

同名输出文件会直接覆盖；如果旧文件正被 Word/WPS/Excel 占用，需要关闭后重试。

触发生成时默认会先清空本次产物对应的输出目录：邮件模式清空 `outputs/邮件/`；SVN 模式清空 `outputs/测试用例评审/` 和 `outputs/会议纪要/`。`--validate-only` 不会清空目录。

## 数据源
数据文件放在 `inputs/requirement_data/*.xlsx`。先从 Ones 按迭代筛选需求并导出 xlsx，再手动补充 `代办事项N@责任人` 列；整理完成后放入该目录。文件名必须带迭代关键字，例如 `SP7`、`SP8`，邮件生成会按该关键字筛选数据文件。

| 字段 | 用途 |
|---|---|
| `ID` | 需求编号；SVN 文件按每行需求拆分 |
| `标题` | 需求标题；用于文档内容和文件名 |
| `测试` | 测试负责人；用于邮件表格、Excel 发起人/主持人 |
| `测试用例链接` | 写入 SVN Excel 的测试用例链接 |
| `计划完成日期` | 自动取前一天作为评审时间/计划解决日期 |
| `前端开发` / `后端开发` / `创建者` | 参会人员来源 |
| `所属迭代` | 邮件按该字段生成，邮件正文和文件名里的 `{iteration}` 也只取该字段；`--data-file-keyword SPx` 只用于筛选文件 |
| `代办事项N@责任人` | 任意数量的代办列，如 `代办事项1@责任人`、`代办事项4@责任人`；脚本按 N 自动展开 |

生成前会自动校验数据。校验失败时不会生成文件；校验警告会继续输出，例如 `测试用例链接` 为空、未检测到代办列、代办列名格式不符合 `代办事项N@责任人`。

## 字段规则
- SVN Excel 只填 `测试用例评审记录` 页指定字段；封皮、签名页、嵌入对象、图片和 13 项检查项结构保持模板原样。
- SVN Word 的 `记录人员` 默认取当前需求行 `测试` 字段；CLI 传入 `--recorder` 时可覆盖。
- SVN Word 的 `会议内容` 写入当前需求所有 `代办事项N@责任人` 原文；没有代办事项时填 `无`。
- SVN Word 的 `改进或遗留工作项` 按代办事项拆行，责任人从 `@姓名` 提取；没有代办事项时表格填 `无`。
- 邮件 Word 的需求表为 `编号 / 标题 / 测试 / 评审结果 / 备注#`，评审结果固定为 `通过`。
- 邮件 Word 的 `1.`、`2.`、`3.` 章节标题加粗；`3.评审记录/代办事项` 下的需求 ID 和需求标题不加粗，字号小于章节标题。
- 邮件 Word 的代办列表保留 `@责任人` 原文。

## 配置
| 文件 | 作用 |
|---|---|
| `config/project.yaml` | 产品名、迭代关键字提取规则等项目级配置 |
| `config/templates.yaml` | SVN Excel/Word 模板和邮件参考图文件名 |
| `config/filename_templates.yaml` | 输出文件名模板、输出子目录 |
| `config/content_rules.yaml` | 邮件文案、会议地点、时间策略、参会人员来源、todo 默认值、Excel 字段来源 |

文件名占位符：`{product}`、`{title}`、`{iteration}`、`{version}`、`{req_id}`。所有占位符会自动替换 Windows 非法文件名字符；标题默认截断到 60 字符。

注意：Excel 封面 D3 的 `版本：A1` 是 SOP 表单修订号，不等于输出文件名里的 `{version}`。生成 SVN 文件前必须向用户确认 `{version}`。
如果用户输入的是纯数字版本（如 `1.0.2`），脚本会自动规范化为 `v1.0.2`；已包含前缀的版本号（如 `v1.0.2`、`SP7`、`CPV-SPx-v1.2`）保持原样。

配置文件是默认值的单一来源。脚本启动时会读取 YAML；配置缺失或格式错误会直接失败，避免 Python 中旧默认值与 YAML 不一致。

## 环境
推荐使用项目虚拟环境：

```bash
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 约束
- SVN 文件必须按需求生成，不得把多条需求合并到同一份 SVN 归档文件。
- 邮件 Word 必须按单个迭代生成，不得跨迭代。
- 模板封皮、签名页、13 项检查项结构不得删除或重排。
- `@责任人` 原文必须保留在邮件和会议内容中。
