import random
import logging
import aiosqlite

logger = logging.getLogger(__name__)


# ── 步长 & Snap ──────────────────────────────────────────────────────────────

def get_step(item_type: str, P: float) -> int:
    if item_type == "doc_pages":
        if P <= 10:   return 1
        elif P <= 50: return 5
        else:         return 10
    elif item_type == "code_loc":
        if P < 100:    return 10
        elif P < 500:  return 50
        elif P < 2000: return 100
        else:          return 500
    return 1  # workload


def snap(v: float, step: int) -> int:
    return max(step, round(v / step) * step)


# ── 4人估算 ──────────────────────────────────────────────────────────────────

def estimate_4(P: float, item_type: str) -> list[int]:
    """生成 4 人估算值，偏差 ≤ 20%，至少 2 个不同值。"""
    step = get_step(item_type, P)

    # 小值保护（文档规模 P < 10，步长已为 1）
    if P < 10 and item_type == "doc_pages":
        avg_target = P * random.uniform(0.95, 1.05)
        base = max(1, round(avg_target))
        return [max(1, base + d) for d in (-2, -1, 1, 2)]

    # 带重试的主路径，确保偏差 ≤ 20%
    for _ in range(50):
        avg_target = P * random.uniform(0.95, 1.05)
        raw = [random.uniform(-0.15, 0.15) for _ in range(4)]
        mean_raw = sum(raw) / 4
        offsets = [r - mean_raw for r in raw]
        Vi = [snap(avg_target * (1 + o), step) for o in offsets]

        if len(set(Vi)) == 1:
            Vi[0] = max(step, Vi[0] - step)
            Vi[-1] = Vi[-1] + step

        avg_vi = sum(Vi) / 4
        if avg_vi > 0 and all(abs(v - avg_vi) / avg_vi <= 0.20 for v in Vi):
            return Vi

    return Vi  # 兜底：返回最后一次结果


# ── 工作量推导 ───────────────────────────────────────────────────────────────

def derive_workload_target(P_ref, N: int, D: int) -> tuple[int, int]:
    """
    返回 (Tm, EM_target)
    Tm: 每人每天工时，clamp [1, 6]
    EM_target = N × D × Tm
    """
    N = max(1, N)
    D = max(1, D)
    if P_ref and P_ref > 0:
        Tm = max(1, min(6, round(P_ref / N / D)))
    else:
        Tm = 4
    return Tm, N * D * Tm


# ── 辅助函数 ─────────────────────────────────────────────────────────────────

def _count_team_n(resources: str | None, team_members: str | None, context: str = "") -> int:
    if not resources:
        if context:
            logger.warning("N fallback=1 for %s: no resources", context)
        return 1
    team_set = {n.strip() for n in team_members.split(",")} if team_members else set()
    resource_list = [n.strip() for n in resources.split(",") if n.strip()]
    if team_set:
        in_team = [n for n in resource_list if n in team_set]
        N = len(in_team) if in_team else len(resource_list)
    else:
        N = len(resource_list)
    if N == 0:
        logger.warning("N fallback=1 for %s: empty intersection", context)
    return max(1, N)


def _best_match(task_name: str, candidates: list[dict], name_key: str) -> dict | None:
    """任务名中包含候选项名的最长匹配。"""
    matches = [c for c in candidates if c.get(name_key) and c[name_key] in task_name]
    if not matches:
        return None
    best = max(matches, key=lambda c: len(c[name_key]))
    if len(matches) > 1:
        logger.warning("Multiple matches for '%s': picked '%s'", task_name, best[name_key])
    return best


# ── Session 生成主函数 ────────────────────────────────────────────────────────

async def generate_estimates(db: aiosqlite.Connection, project_id: int) -> dict:
    # 读取项目
    async with db.execute("SELECT * FROM projects WHERE id=?", (project_id,)) as cur:
        project = dict(await cur.fetchone())
    team_members = project.get("team_members") or ""
    team_set = {n.strip() for n in team_members.split(",")} if team_members else set()

    # 创建 Session
    async with db.execute(
        "INSERT INTO estimate_sessions (project_id) VALUES (?)", (project_id,)
    ) as cur:
        session_id = cur.lastrowid

    # 读取所有数据
    async with db.execute(
        "SELECT * FROM config_items WHERE project_id=? ORDER BY seq, id", (project_id,)
    ) as cur:
        all_items = [dict(r) for r in await cur.fetchall()]
    doc_items = [i for i in all_items if i.get("item_type") == "文档"]

    async with db.execute(
        "SELECT * FROM code_modules WHERE project_id=? ORDER BY seq, id", (project_id,)
    ) as cur:
        code_modules = [dict(r) for r in await cur.fetchall()]

    async with db.execute(
        "SELECT * FROM wbs_tasks WHERE project_id=? ORDER BY id", (project_id,)
    ) as cur:
        all_tasks = [dict(r) for r in await cur.fetchall()]

    eng_tasks = [t for t in all_tasks if t.get("task_type") == "engineering"]
    phase_map = {t["wbs"]: t["name"] for t in all_tasks if t.get("task_type") == "phase"}

    # 第一阶段识别：WBS 最小的 phase 任务，其子任务及文档均不参与估算
    phase_wbs_list = [t["wbs"] for t in all_tasks if t.get("task_type") == "phase" and t.get("wbs")]
    first_phase_wbs = min(phase_wbs_list, key=lambda w: [int(x) for x in w.split(".") if x.isdigit()]) if phase_wbs_list else None
    first_phase_prefix = (first_phase_wbs + ".") if first_phase_wbs else None

    def _in_first_phase(task: dict) -> bool:
        if not first_phase_prefix:
            return False
        return (task.get("wbs") or "").startswith(first_phase_prefix)

    # 第一阶段文档（phase_wbs 字段与第一阶段WBS一致）不参与估算
    doc_items = [i for i in doc_items
                 if not (first_phase_wbs and str(i.get("phase_wbs") or "").strip() == first_phase_wbs)]

    # ── 工程任务匹配 ────────────────────────────────────────────────────────
    doc_map: dict[int, dict] = {}      # config_item_id → task（记录匹配关系供参考）
    code_map: dict[int, dict] = {}     # module_id → task
    sheet1_wl_tasks: list[dict] = []   # Sheet1 工作量区：所有非代码工程任务（按 WBS 顺序）

    for task in eng_tasks:
        if _in_first_phase(task):
            continue  # 第一阶段工程任务不进 Sheet1 工作量区

        code_match = _best_match(task["name"], code_modules, "name")
        if code_match and code_match["id"] not in code_map:
            code_map[code_match["id"]] = task
            continue  # 代码实现任务归 Sheet2，不进 Sheet1

        doc_match = _best_match(task["name"], doc_items, "name")
        if doc_match and doc_match["id"] not in doc_map:
            doc_map[doc_match["id"]] = task  # 记录匹配，供规模行写 wbs_task_id

        sheet1_wl_tasks.append(task)  # 文档编写 & 其他工程任务都进 Sheet1 工作量区

    # ── Sheet 1：文档规模行（仅规模，无工作量） ─────────────────────────────
    for item in doc_items:
        actual_pages = item.get("actual_pages")
        matched = doc_map.get(item["id"])

        if actual_pages:
            v1, v2, v3, v4 = estimate_4(float(actual_pages), "doc_pages")
        else:
            v1 = v2 = v3 = v4 = None

        wbs_task_id = matched["id"] if matched else None

        await db.execute(
            """INSERT INTO sheet1_estimates
               (session_id, config_item_id, wbs_task_id, item_name, actual_pages,
                v1, v2, v3, v4, em_target, wl1, wl2, wl3, wl4)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (session_id, item["id"], wbs_task_id, item["name"], actual_pages,
             v1, v2, v3, v4, None, None, None, None, None),
        )

    # ── Sheet 1：工程活动工作量行（文档编写 + 其他工程，按 WBS 顺序） ────────
    for task in sheet1_wl_tasks:
        N = _count_team_n(task.get("resources"), team_members, f"eng:{task['name']}")
        D = task.get("duration_days") or 1
        _, em_target = derive_workload_target(task.get("workload"), N, D)
        wl1, wl2, wl3, wl4 = estimate_4(float(em_target), "workload")

        await db.execute(
            """INSERT INTO sheet1_estimates
               (session_id, config_item_id, wbs_task_id, item_name,
                em_target, wl1, wl2, wl3, wl4)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (session_id, None, task["id"], task["name"],
             em_target, wl1, wl2, wl3, wl4),
        )

    # ── Sheet 2：代码规模 + 工作量 ──────────────────────────────────────────
    for mod in code_modules:
        actual_loc = mod.get("actual_loc")
        matched = code_map.get(mod["id"])

        if actual_loc:
            v1, v2, v3, v4 = estimate_4(float(actual_loc), "code_loc")
        else:
            v1 = v2 = v3 = v4 = None

        em_target = wl1 = wl2 = wl3 = wl4 = None
        wbs_task_id = None

        if matched:
            wbs_task_id = matched["id"]
            N = _count_team_n(matched.get("resources"), team_members,
                              f"code:{mod['name']}")
            D = matched.get("duration_days") or 1
            _, em_target = derive_workload_target(matched.get("workload"), N, D)
            wl1, wl2, wl3, wl4 = estimate_4(float(em_target), "workload")

        await db.execute(
            """INSERT INTO code_scale_estimates
               (session_id, module_id, wbs_task_id, module_name, actual_loc,
                v1, v2, v3, v4, em_target, wl1, wl2, wl3, wl4)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (session_id, mod["id"], wbs_task_id, mod["name"] or "", actual_loc,
             v1, v2, v3, v4, em_target, wl1, wl2, wl3, wl4),
        )

    # ── Activity 快照（CM / QA / mgmt）──────────────────────────────────────
    for task in all_tasks:
        if task.get("task_type") not in ("cm", "qa", "mgmt"):
            continue
        if _in_first_phase(task):
            continue  # 第一阶段活动不参与估算
        wbs = task.get("wbs") or ""
        phase_name = phase_map.get(wbs.split(".")[0]) if wbs else None
        await db.execute(
            """INSERT INTO activity_workload_snapshots
               (session_id, task_id, activity_type, phase_name, activity_name, workload)
               VALUES (?,?,?,?,?,?)""",
            (session_id, task["id"], task["task_type"], phase_name,
             task["name"], task.get("workload")),
        )

    # ── Review 快照 ─────────────────────────────────────────────────────────
    for task in all_tasks:
        if task.get("task_type") != "review":
            continue
        if _in_first_phase(task):
            continue  # 第一阶段评审不参与估算
        resources = task.get("resources") or ""
        resource_list = [n.strip() for n in resources.split(",") if n.strip()]
        if team_set:
            in_team = [n for n in resource_list if n in team_set]
            participants = len(in_team) if in_team else (len(resource_list) or 1)
        else:
            participants = len(resource_list) or 1
        await db.execute(
            """INSERT INTO review_workload_snapshots
               (session_id, task_id, activity_name, participants, total_hours)
               VALUES (?,?,?,?,?)""",
            (session_id, task["id"], task["name"], participants, task.get("workload")),
        )

    await db.commit()

    # 返回统计
    counts = {}
    for table, key in [
        ("sheet1_estimates", "sheet1_rows"),
        ("code_scale_estimates", "code_scale_rows"),
        ("activity_workload_snapshots", "activity_rows"),
        ("review_workload_snapshots", "review_rows"),
    ]:
        async with db.execute(
            f"SELECT COUNT(*) FROM {table} WHERE session_id=?", (session_id,)
        ) as cur:
            counts[key] = (await cur.fetchone())[0]

    return {"session_id": session_id, **counts}
