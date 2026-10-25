
## 项目上下文（每次新会话必读）

### 这是什么项目
EstimateM：项目配置项规模及工作量估算记录生成系统。给定项目完成后的实际数据（文档页数、代码行数、进度计划工作量），逆向生成符合公司规范的「4人独立估算 + 汇总 + 结果报告」10-Sheet Excel文件。
技术栈：Python + FastAPI + SQLite (aiosqlite) + Alpine.js + openpyxl

### 关键文档路径
- **设计文档（APPROVED）**：`~/.gstack/projects/EstimateM/gaoyiman-unknown-design-20260625-150004.md`
- **开发计划与进度**：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/DEV_PLAN.md`
- **样例估算文件（格式参照）**：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/规模及工作量估计表.xlsx`
- **样例输入文件（字段参照）**：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/项目相关信息表.xlsx`
- **Excel 布局参考（excel_builder.py 必读）**：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/excel_layout.md`

### 当前阶段（更新于 2026-06-29）
**Phase 5 已完成，全流程可用。**

### Windows 离线交付（2026-10-26）

- 已按批准方案完成本地 Alpine.js 3.14.1、仅监听 `127.0.0.1:8001`、关闭 `/docs` 与 `/redoc`、Windows 启动脚本及交付工具。
- 正式操作说明：`DEPLOY_WINDOWS_OFFLINE.md`；验证范围和未完成项：`DEPLOYMENT_VERIFICATION.md`。
- Windows 基线：标准 CPython 3.13.17 x64，使用 `requirements-win-py313.lock.txt`（含 wheel SHA-256）；锁文件是按 Windows 目标跨平台解析的候选基线，仍需 Windows 实机验收。
- Mac 继续系统 Python + `python3 run.py`，不安装 Windows 锁文件；Windows 独立 `.venv`，双击 `start_windows.cmd`。不变更估算算法或导入导出规则。
- 直接沿用 `项目相关信息表.xlsx`：使用人员复制样例后修改，不新增空白模板。输出参考 xlsx 不是运行时必需资源。
- 各电脑全新建库，不交付现有 DB/历史导出；只打包明确允许列表。wheelhouse 与 Python 安装程序由部署人员在联网 Windows 准备，不放入 Git。
- 不依赖开发分析脚本中的作者电脑绝对路径。仓库仍保持私有，未来公开前须单独检查样例、文档及历史并取得公开授权；不新增 Actions。
- 下一步：部署人员执行 Windows 完整离线安装、双击启动、浏览器与 Excel/WPS 验收，记录结果后才能标记 Windows 交付验收完成。

已确认：
- ✅ 输入 Excel 扩展字段（项目信息 19 列、进度计划新增工期/工作量/资源名称、配置项清单新增类型/实际页数、新增代码模块 Sheet）
- ✅ 输出 10-Sheet Excel 结构（参照样例文件，动态行数全量 openpyxl 生成）
- ✅ 数据库 Schema（projects / config_items / code_modules / wbs_tasks / estimate_sessions / sheet1_estimates / code_scale_estimates / activity_workload_snapshots / review_workload_snapshots / generated_files）
- ✅ 4人估算算法（步长规则 + snap + 偏差约束）
- ✅ 工作量推导算法（Tm = P/N/D 取整 clamp[1,6]，EM_target = N×D×Tm）
- ✅ UI 设计（5个Tab：项目信息 / 配置项管理 / 进度计划 / 估算结果 / 生成）
- ✅ 任务类型识别规则（7类，含 engineering 叶节点识别）
- ✅ Phase 1：database.py / run.py / main.py — 服务可启动，10张表结构就绪
- ✅ Phase 2：importer.py / task_classifier.py / 三个路由 — Excel导入正常，任务类型识别全通过
- ✅ Phase 3：estimator.py / estimates.py — 算法单元测试全通过，Session 生成/删除正常
- ✅ Phase 4：excel_builder.py / export.py — 10-Sheet Excel 生成正常，GET /api/sessions/{id}/export 返回 xlsx 文件
- ✅ Phase 5：index.html — Alpine.js 5 Tab UI，从 Excel 一键导入建项目，全流程可用

**下一步：按需修复 / 用户验收测试**

**UI 已知约定（新会话必知）**：
- 顶部「从 Excel 导入」按钮：自动创建项目 + 导入 Excel（读取项目名/编号等覆盖占位值），导入后跳转 Tab 3
- 顶部「+ 手动新建」按钮：手动填写项目信息，之后在 Tab 3 可再次导入 Excel
- Tab 2 配置项管理：**不显示工作量列**（config_items.workload 字段始终为空，workload 目标由 Session 估算生成，不在此维护）
- Tab 4 估算结果需先在 Tab 5 生成 Session，才能在下拉框中选择

**注意**：端口使用 **8001**（8000 被 ConfigM 占用）

### 项目结构（目标）

```
EstimateM/
├── run.py                        # 启动入口：python3 run.py
├── estimatem.db                  # SQLite 数据库
├── generated/                    # 生成的 Excel 文件
│   └── {project_id}/
├── app/
│   ├── main.py                   # FastAPI app + 路由注册
│   ├── database.py               # aiosqlite 连接 + DDL 初始化
│   ├── models.py                 # Pydantic 请求/响应模型
│   ├── routers/
│   │   ├── projects.py           # CRUD 项目信息
│   │   ├── config_items.py       # 配置项 + 代码模块 CRUD
│   │   ├── wbs_tasks.py          # 进度计划（含 task_type 识别）
│   │   ├── estimates.py          # 生成估算、查看/编辑估算结果
│   │   └── export.py             # 导出 Excel（从 DB 数据生成）
│   ├── services/
│   │   ├── importer.py           # 解析项目相关信息表.xlsx → DB
│   │   ├── task_classifier.py    # WBS 任务类型自动识别
│   │   ├── estimator.py          # 4人估算算法 + 工作量推导
│   │   └── excel_builder.py      # openpyxl 生成 10-Sheet Excel
│   └── static/
│       └── index.html            # Alpine.js 单页应用（5 Tab）
└── requirements.txt
```

### 关键规则（新会话必知）

**启动方式**
```bash
python3 run.py   # 系统 Python，不使用 venv
```

**任务类型识别（task_classifier.py）**

| 识别条件 | task_type |
|----------|-----------|
| WBS 级别 = 1 | `phase` |
| 祖先节点名 = 配置管理活动，且为叶节点 | `cm` |
| 祖先节点名 = 质量保证活动，且为叶节点 | `qa` |
| 祖先节点名 = 评审活动，且为叶节点 | `review` |
| 父节点名 ∈ {项目监控活动, 需求管理活动, 策划活动}，且为叶节点（直接子任务） | `mgmt` |
| 祖先节点名 = 软件工程活动，且为叶节点 | `engineering` |
| 以上均不匹配 | `other` |

父节点查找：WBS = "1.3.2" → 父 WBS = "1.3"（最长前缀匹配，后跟"."）
叶节点判断：该行在 WBS 列中不存在任何以其 WBS+"." 开头的其他行

**engineering 任务子分类（excel_builder.py）**

| 子类 | 识别 | Excel 落点 |
|------|------|------------|
| 代码实现 | 任务名包含某 `code_modules.name` | Sheet 2 工作量区 |
| 文档编写 | 任务名包含某 `config_items.name`（类型=文档） | Sheet 1 工作量区（任务名作为估计对象，**不重复**出现在规模区） |
| 其他工程（测试等） | 以上均不匹配 | Sheet 1 工作量区 |

**注意**：Sheet 1 规模区（"页"）仅列文档配置项；工作量区（"人时"）仅列 WBS 工程任务（文档编写任务 + 其他工程任务），两者不重复，不存在空的工作量行。

**4人估算算法步长规则（estimator.py）**

| 类型 | 区间 | 步长 |
|------|------|------|
| 文档规模（页） | P ≤ 10 | 1 |
| 文档规模（页） | 10 < P ≤ 50 | 5 |
| 文档规模（页） | P > 50 | 10 |
| 代码规模（LOC） | P < 100 | 10 |
| 代码规模（LOC） | 100 ≤ P < 500 | 50 |
| 代码规模（LOC） | 500 ≤ P < 2000 | 100 |
| 代码规模（LOC） | P ≥ 2000 | 500 |
| 工作量（人时） | — | 1（不 snap） |

各人偏差 ≤ 20%；4值全相同时微调首尾 ±1 步

**工作量推导公式（estimator.py）**
```
Tm = clamp(round(P_ref / N / D), 1, 6)   # P_ref 缺失时默认 4
EM_target = N × D × Tm
```
N = 资源名称中属于项目组成员的人数（与 projects.team_members 取交集）

**估算结果存储（估算结果可在 Tab 4 编辑，导出 Excel 从 DB 值生成不重算）**
- A组（4人估算）：`sheet1_estimates`（Sheet 1 所有行：文档规模+工程工作量）、`code_scale_estimates`，V1-V4 可编辑
- B组（单人实际值快照）：`activity_workload_snapshots`（cm/qa/mgmt）、`review_workload_snapshots`（含 participants 和 total_hours，无 duration_days），workload 可编辑

**重导入规则**：重新导入 Excel 时，全量覆盖 wbs_tasks（含 task_type），不保留手动修改

**Excel 输出格式**
- 严格对齐：列结构、合并单元格区域、跨 Sheet 公式、表头文字、行高列宽
- 动态扩展：文档/代码行数随实际数量变化，汇总表公式范围同步调整
- 参照文件：`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/规模及工作量估计表.xlsx`

**项目信息关键字段**
- `projects.code` → 项目编号前缀 `ZDPX{code}`
- `projects.estimators` → JSON `{doc:[4人名], code:[4人名], cm, qa, mgmt, review}`
- `projects.team_members` → 逗号分隔，用于过滤资源名称中的项目组内成员

### 里程碑更新规则
**每完成一个功能模块后，必须立即同步更新以下两个文件，再继续下一模块：**

1. **本文件（CLAUDE.md）**
   - 更新「当前阶段」的描述和日期
   - 打 ✅ 标记刚完成的项
   - 更新「下一步」为下一阶段名称
   - 补充编码过程中发现的新规则或约定

2. **DEV_PLAN.md**（`/Users/gaoyiman/Desktop/Code/5000B/EstimateM/DEV_PLAN.md`）
   - 勾选已通过的验收 checkbox `[ ]` → `[x]`
   - 将对应 Phase 状态从 `⏳ 待开始` 或 `🔨 进行中` 改为 `✅ 通过`
   - 更新文件顶部「最后更新」日期

---

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
