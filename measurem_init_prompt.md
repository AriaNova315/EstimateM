我要在当前目录里创建一个新项目，MeasureM，用于生成"项目测量与监控记录"系列 Excel 文件。

## 项目背景

这是一套工具集的第三个系统，前两个系统是：
- ConfigM（配置管理记录生成），位于同级目录 ../ConfigM/
- EstimateM（规模及工作量估算记录生成），位于同级目录 ../EstimateM/

MeasureM 的定位：项目**结束后**，对照实际数据补充生成测量与监控阶段所需的全套记录文件。
目标是生成看起来足够真实、逻辑自洽的记录，而非实时采集数据。

## 架构决策（已确定，不再讨论）

每个系统完全独立：
- MeasureM 维护自己的 SQLite 数据库 measurem.db
- 用户向 MeasureM 导入同一份项目 Excel（项目相关信息表.xlsx），MeasureM 只解析自己所需的 Sheet 和列
- 不读取 ConfigM / EstimateM 的数据库，不依赖其他系统服务
- Excel 文件是数据的权威来源，各系统从同一份完整 Excel 各取所需
- 后续如需扩展，可在项目相关信息表.xlsx 中新增列或 Sheet，不影响其他系统

## 技术栈（已确定，不再讨论）

与 ConfigM / EstimateM 完全一致：
- Python + FastAPI + SQLite（aiosqlite）+ Alpine.js + openpyxl
- 启动方式：`python3 run.py`，使用系统 Python，不使用 venv
- 端口：**8002**（8000 被 ConfigM 占用，8001 被 EstimateM 占用）
- 前端：单页 Alpine.js 应用，**UI 风格与 EstimateM 保持一致**

## 核心功能范围（初步，待需求讨论细化）

MeasureM 需要能够生成以下记录内容（每个项目阶段各生成一套文件）：

1. **工时记录及统计** —— 项目组成员在本阶段内各自的工时分配与汇总
2. **任务定期进度测量** —— 各 WBS 任务在测量周期内的计划 vs 实际进度数据
3. **测量分析报告与进展报告** —— 基于测量数据形成的分析结论
4. **例会单** —— 阶段内各次例会的会议记录单

每个阶段对应一份输出 Excel 文件（含多个 Sheet），文件数量随阶段数动态变化。

## 输出方式（已确定，与 EstimateM 一致）

- 流程：导入数据 → UI 中查看/调整生成参数 → 点击下载按钮 → 实时生成 Excel 并下载
- 格式：openpyxl 生成多 Sheet Excel，格式参照样例文件
- 命名规则：待需求讨论确定

## 现有 Excel 中可复用的字段

项目相关信息表.xlsx 目前包含以下 Sheet（EstimateM 已扩展的版本）：

1. **项目信息**：项目名称、项目代号、项目组长、配置管理员、质量保证人员、估算轮次、各类估算人员、团队成员名单
2. **项目进度计划**：WBS、任务名称、计划开始/完成时间、关联配置项、工期（天）、工作量（人时）、资源名称
3. **配置项清单**：序号、配置项名称、标识、版本、类型、实际页数、所属阶段 WBS
4. **代码模块**：序号、模块名称、实际代码行数

MeasureM 主要复用：项目信息、项目进度计划（WBS 结构、阶段划分、任务/资源信息）。
后续如需录入工时基准、测量频率等信息，可讨论是否在 Excel 中新增字段或在 UI 中录入。

## 样例文件（输出格式的权威参照）

已有一份现成的输出样例文件，需先读取并分析其结构，再开展需求讨论：

**样例文件路径**：`/Users/gaoyiman/Desktop/Code/5000B/高蕾/02.需求分析阶段.xlsx`

该文件对应"需求分析"阶段的完整一套测量与监控记录，包含以下 Sheet：

| Sheet 名称 | 内容说明 |
|-----------|---------|
| `WorkReport20250804` | 工时记录表：姓名 × WBS × 任务名称 × 计划/实际时间 × 完成百分比 × 工时合计，并按日展开动态列（每日各人各任务实际工时） |
| `ProgressTracking-ok` | 任务进度跟踪表：WBS 层级 × 任务信息 × 各测量周期列（计划应完成工作量、实际完成百分比、实际已完成工作量、实际工时、进度偏差、工作量偏差） |
| `002测量与分析报告（阶段）` | 阶段内定期测量分析报告：包含项目信息表头、规模/工作量/进度/需求各维度的计划 vs 实际偏差分析、异常分析结论、拟制人员 |
| `002进展报告（阶段）` | 阶段内定期进展报告：项目信息表头、进展情况（任务完成叙述、偏差统计）、风险/资源/培训/问题跟踪、工作总结与下期安排、文档清单 |
| `LH004项目例会单` | 第4次例会记录单：时间/地点/主持人/记录员/参加人员/主要内容（含多条议题） |
| `LH005项目例会单` | 第5次例会记录单：同上格式 |
| `003测量与分析报告（里程碑）` | 里程碑测量分析报告：覆盖整个阶段周期，格式与阶段报告类似但内容更完整 |
| `005进展报告（里程碑）` | 里程碑进展报告：覆盖系统分析与设计和需求分析两个阶段，包含QA/CM工作汇总，比阶段进展报告内容更全面 |

> **重要**：在触发 /office-hours 之前，请先用 openpyxl 或 Python 读取并分析该样例文件的详细结构（列名、合并单元格、行高、字体、边框等），以便在需求讨论中能够精确描述格式要求，并在后续编码时直接参照。

## CLAUDE.md 要求

请参考以下结构创建 CLAUDE.md，与 ConfigM / EstimateM 保持同一风格：

1. 项目上下文（每次新会话必读）：项目定位、技术栈、关键文档路径
2. 当前阶段（已完成 / 下一步）
3. 关键文件清单与项目结构
4. 已知规则和约定
5. 里程碑更新规则
6. Skill routing（内容如下，直接复制写入 CLAUDE.md）

## Skill routing

When the user's request matches an available skill, invoke it via the Skill tool. When in doubt, invoke the skill.

Key routing rules:
- Product ideas/brainstorming → invoke /office-hours
- Strategy/scope → invoke /plan-ceo-review
- Architecture → invoke /plan-eng-review
- Design system/plan review → invoke /design-consultation or /plan-design-review
- Full review pipeline → invoke /autoplan
- Bugs/errors → invoke /investigate
- QA/testing site behavior → invoke /qa or /qa-only
- Code review/diff check → invoke /review
- Visual polish → invoke /design-review
- Ship/deploy/PR → invoke /ship or /land-and-deploy
- Save progress → invoke /context-save
- Resume context → invoke /context-restore

## 现在请执行

**第一步：先读取并分析样例文件**

用 Python + openpyxl 读取 `/Users/gaoyiman/Desktop/Code/5000B/高蕾/02.需求分析阶段.xlsx`，提取并输出每个 Sheet 的：
- 列结构（列名、数据类型、合并单元格范围）
- 关键行高（表头行、数据行）
- 字体与边框规范
- 动态列的规律（如 WorkReport 的日期列、ProgressTracking 的周期列）

**第二步：触发 /office-hours**，围绕以下问题展开需求澄清，下面问题仅是建议，你也可以根据你的理解进行优化和补充（格式问题已由样例文件回答，重点讨论数据生成逻辑与操作流程）：

**关于数据生成逻辑**
- 工时记录（WorkReport）：每人每日工时如何从进度计划中推算？是基于任务工作量 + 资源名称 + 工期按天平摊，还是有其他规则？每天总工时上限是多少？
- 进度跟踪（ProgressTracking）：每个测量周期的"实际完成百分比"如何生成？是基于计划值加随机扰动，还是有固定规则（如线性递增）？偏差率的合理范围是多少？
- 测量分析报告：报告中的叙述性文字（分析结论、异常原因）是模板填充还是需要用户手动输入？
- 进展报告：其中的问题跟踪、风险条目是固定模板还是需要用户维护？
- 例会单：例会日期和议题内容如何确定？是用户逐条录入，还是有默认模板？

**关于阶段与报告类型**
- 每个阶段内，定期报告（阶段类）生成几份？是固定2份（样例中的 002 和 003 对应一次阶段、一次里程碑），还是用户指定测量次数？
- "阶段"报告和"里程碑"报告的区别在数据上如何体现（时间范围不同？内容深度不同）？

**关于操作流程**
- 用户完整操作路径：导入 Excel → 在 UI 中补充哪些参数（阶段时间范围？测量日期？例会日期列表？） → 触发生成 → 下载
- 生成单个阶段文件还是一次生成所有阶段？
- 是否需要保存生成历史（类似 EstimateM 的 Session 概念）？

**关于文件命名**
- 样例文件名为 `02.需求分析阶段.xlsx`，命名规则是"序号.阶段名称.xlsx"？
- 文件内 Sheet 的编号（如 `002`、`003`、`LH004`）代表什么含义？是否需要用户指定或自动生成？

**其他你分析需要进一步确认的问题**
