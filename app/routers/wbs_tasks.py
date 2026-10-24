from fastapi import APIRouter, HTTPException

from app.database import get_db
from app.models import WbsTaskUpdate, WbsTaskResponse

router = APIRouter(tags=["wbs_tasks"])

_VALID_TYPES = {'phase', 'cm', 'qa', 'review', 'mgmt', 'engineering', 'other'}


@router.get("/api/projects/{project_id}/tasks", response_model=list[WbsTaskResponse])
async def list_tasks(project_id: int):
    db = await get_db()
    sql = """
        WITH ls AS (
            SELECT id FROM estimate_sessions WHERE project_id=? ORDER BY id DESC LIMIT 1
        )
        SELECT t.*,
            COALESCE(s1.em_target, aws.workload, rws.total_hours) AS latest_est_workload
        FROM wbs_tasks t
        LEFT JOIN sheet1_estimates s1
            ON s1.wbs_task_id = t.id
           AND s1.session_id = (SELECT id FROM ls)
           AND s1.config_item_id IS NULL
        LEFT JOIN activity_workload_snapshots aws
            ON aws.task_id = t.id
           AND aws.session_id = (SELECT id FROM ls)
        LEFT JOIN review_workload_snapshots rws
            ON rws.task_id = t.id
           AND rws.session_id = (SELECT id FROM ls)
        WHERE t.project_id=?
        ORDER BY t.id
    """
    async with db.execute(sql, (project_id, project_id)) as cur:
        rows = await cur.fetchall()
    return [dict(r) for r in rows]


@router.put("/api/tasks/{task_id}", response_model=WbsTaskResponse)
async def update_task(task_id: int, data: WbsTaskUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    if "task_type" in fields and fields["task_type"] not in _VALID_TYPES:
        raise HTTPException(400, f"Invalid task_type: {fields['task_type']}")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE wbs_tasks SET {set_clause} WHERE id=?",
        list(fields.values()) + [task_id],
    )
    await db.commit()
    async with db.execute("SELECT * FROM wbs_tasks WHERE id=?", (task_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Task not found")
    return dict(row)
