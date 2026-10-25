# EstimateM 开发计划与进度

> 最后更新：2026-10-26（Phase 5 完成；Windows 离线交付实现完成，实机验收待完成）
> 原则：每阶段通过验收标准后再进入下一阶段

---

## 总览

| 阶段 | 名称 | 状态 | 核心产出 |
|------|------|------|---------|
| Phase 1 | 基础骨架 | ✅ 通过 | database.py / run.py / main.py |
| Phase 2 | 数据导入与 CRUD | ✅ 通过 | importer.py / task_classifier.py / 三个路由 |
| Phase 3 | 估算算法与 Session | ✅ 通过 | estimator.py / estimates.py |
| Phase 4 | Excel 导出 | ✅ 通过 | excel_builder.py / export.py |
| Phase 5 | 前端 UI | ✅ 通过 | static/index.html |

状态说明：⏳ 待开始 / 🔨 进行中 / ✅ 通过 / ❌ 阻塞

---

## Phase 1 — 基础骨架

**目标**：服务可启动，数据库结构就绪

### 实现文件
- `app/database.py` — aiosqlite 连接池 + 所有 DDL（10 张表）
- `run.py` — uvicorn 启动入口
- `app/main.py` — FastAPI app，路由注册占位
- `requirements.txt`

### 数据库表清单
- [x] projects
- [x] config_items
- [x] code_modules
- [x] wbs_tasks
- [x] estimate_sessions
- [x] sheet1_estimates
- [x] code_scale_estimates
- [x] activity_workload_snapshots
- [x] review_workload_snapshots
- [x] generated_files

### 验收标准
- [x] `python3 run.py` 启动无报错，监听 8001 端口（8000 被 ConfigM 占用）
- [x] `estimatem.db` 自动创建，10 张表结构正确
- [x] `GET /health` 返回 200

### 状态：✅ 通过

---

## Phase 2 — 数据导入与 CRUD

**目标**：样例 Excel 可导入，任务类型自动识别正确

### 实现文件
- `app/services/importer.py` — 解析 5 个 Sheet 写入 DB
- `app/services/task_classifier.py` — WBS 任务类型识别
- `app/routers/projects.py` — 项目 CRUD
- `app/routers/config_items.py` — 配置项 + 代码模块 CRUD
- `app/routers/wbs_tasks.py` — 进度计划 CRUD + 导入接口
- `app/models.py` — Pydantic 请求/响应模型

### task_classifier 用例（必须全通过）
- [x] WBS="1" → `phase`
- [x] WBS="5.1.1" 祖先含"配置管理活动" + 叶节点 → `cm`
- [x] WBS="3.2.1" 祖先含"软件工程活动" + 叶节点 → `engineering`
- [x] WBS="3.2.1.1" 多级嵌套 → `engineering`
- [x] WBS="4.1" 父节点="项目监控活动" + 叶节点 → `mgmt`
- [x] WBS="4.1.1" 父节点非直接 → `other`
- [x] 无匹配 → `other`
- [x] WBS="1.3.10" 父节点是"1.3"而非"1.3.1"（前缀陷阱）
- [x] WBS="1" find_parent() → None

### 验收标准
- [x] 解析样例 `项目相关信息表.xlsx` 写入 DB 无异常（196任务，26配置项）
- [x] task_classifier 所有用例通过
- [x] `POST /api/projects/{id}/import` 后 `GET /api/projects/{id}/tasks` 返回正确 task_type

### 状态：✅ 通过

---

## Phase 3 — 估算算法与 Session 管理

**目标**：估算生成正确，Session 可管理

### 实现文件
- `app/services/estimator.py` — get_step / snap / estimate_4 / derive_workload_target
- `app/routers/estimates.py` — POST 生成估算 / DELETE 删除 session

### estimator 用例（必须全通过）
- [x] get_step("doc_pages", 5) → 1
- [x] get_step("doc_pages", 25) → 5
- [x] get_step("doc_pages", 100) → 10
- [x] get_step("code_loc", 50) → 10
- [x] get_step("code_loc", 200) → 50
- [x] get_step("code_loc", 1000) → 100
- [x] get_step("code_loc", 3000) → 500
- [x] snap(17, 5) → 15
- [x] snap(2, 5) → 5（不低于一个步长）
- [x] estimate_4(P=20, "doc_pages")：max(Vi)/avg - 1 ≤ 0.20
- [x] estimate_4(P=20, "doc_pages")：4值不全相同
- [x] estimate_4(P=5, "doc_pages")：4个不同值（小值保护）
- [x] derive_workload_target(P_ref=40, N=2, D=5) → Tm=4, EM=40
- [x] derive_workload_target(P_ref=None, N=3, D=3) → Tm=4, EM=36
- [x] derive_workload_target(P_ref=100, N=5, D=10) → Tm=2（clamp生效）
- [x] ROUNDUP(EM / N / Tm, 0) == D 恒成立

### 验收标准
- [x] estimator 所有用例通过
- [x] POST 生成后 4 张估算表行数正确（sheet1=20, activity=115, review=9）
- [x] DELETE /api/sessions/{id} 级联删除 → 204；不存在 → 404
- [x] 重复生成创建新 session，旧数据不变

### 状态：✅ 通过

---

## Phase 4 — Excel 导出

**目标**：生成的 xlsx 结构与样例文件对齐

### 实现文件
- `app/services/excel_builder.py` — 10 个 Sheet 完整生成
- `app/routers/export.py` — GET /api/sessions/{id}/export

### Sheet 验收清单
- [x] Sheet 1/2：行数 = 文档数+工程任务数；4人估算值/公式正确
- [x] Sheet 3/4：汇总公式跨 Sheet 引用正确；H/K/M 列正确
- [x] Sheet 5 CM：行数与 cm 任务数一致
- [x] Sheet 6 QA：行数与 qa 任务数一致
- [x] Sheet 7 管理：行数与 mgmt 任务数一致
- [x] Sheet 8 评审：C=participants, D=total_hours/participants, E==C*D；participants=0 不崩溃
- [x] Sheet 9/10：规模/工作量结果报告汇总正确
- [x] 生成文件可被 Excel/WPS 打开无报错

### 状态：✅ 通过

---

## Phase 5 — 前端 UI

**目标**：浏览器完整操作全流程

### 实现文件
- `app/static/index.html` — Alpine.js 5 Tab 单页应用

### Tab 功能清单
- [x] Tab 1 项目信息：显示/编辑项目基本信息
- [x] Tab 2 配置项管理：文档配置项（不含工作量列）+ 代码模块列表，实际页数/LOC 可编辑
- [x] Tab 3 进度计划：WBS 任务列表，显示 task_type badge，支持按名称/WBS/类型过滤，可导入 Excel
- [x] Tab 4 估算结果：V1-V4 可在线编辑；cm/qa/mgmt/review 工作量可编辑
- [x] Tab 5 生成：选择/切换 Session，生成/重新生成，下载 Excel，删除历史 Session

### 验收标准
- [x] 5 Tab 全部可交互，无 JS 错误
- [x] 导入 → 分类 → 生成 → 下载 完整流程跑通
- [x] Session 删除后列表更新

### 额外实现
- [x] 顶部「从 Excel 导入」一键建项目（自动创建占位项目 + 导入 + 读取项目信息覆盖）
- [x] Tab 2 配置项管理去掉工作量列（该字段不由用户维护）

### 状态：✅ 通过

---

## 已知风险与约束

### Windows 单机离线交付

- [x] 方案经用户批准，沿用已有私有 GitHub 仓库，不增加 CI。
- [x] Alpine.js 3.14.1 本地化，保留 MIT 许可证及来源/校验值。
- [x] 服务仅监听本机，移除开发文档页面公网资源依赖。
- [x] 新增 Windows 启动脚本、完整候选依赖锁文件、允许列表源码打包与隔离回归脚本。
- [x] 新增部署说明，沿用原输入样例，不另做空白模板；明确不迁移现有数据。
- [ ] Windows 11 x64 + CPython 3.13.17 完整安装程序及纯离线依赖安装实测。
- [ ] Windows 双击启动、浏览器空缓存断网、Excel/WPS 重算/布局验收。

证据和限制以 `DEPLOYMENT_VERIFICATION.md` 为准；未通过 Windows 实机验收前，不标记离线部署整体验收完成。Mac 启动方式保持 `python3 run.py`。

| 风险 | 处理方案 |
|------|---------|
| engineering 任务名多子串匹配 | 最长匹配优先；无匹配时记 WARNING |
| participants=0 导出除零 | Sheet 8 D列回退写 total_hours |
| N=0 资源为空 | derive_workload_target 中 N=max(1,0)，记 WARNING |
| WBS 前缀误匹配（"1.3.10"→"1.3.1"） | find_parent 用"."终止符精确匹配 |

---

## 关键参考文件

| 文件 | 用途 |
|------|------|
| `excel_layout.md` | Excel builder 列坐标参考，编写 Phase 4 必读 |
| `规模及工作量估计表.xlsx` | 输出样例文件 |
| `项目相关信息表.xlsx` | 输入样例文件 |
| `~/.gstack/projects/EstimateM/gaoyiman-unknown-design-20260625-150004.md` | 完整设计文档 |
| `~/.gstack/projects/EstimateM/gaoyiman-unknown-eng-review-test-plan-20260626-152731.md` | 测试用例清单 |
