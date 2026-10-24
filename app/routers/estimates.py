from fastapi import APIRouter, HTTPException

from app.database import get_db
from app.models import (
    SessionResponse,
    Sheet1EstimateUpdate, Sheet1EstimateResponse,
    CodeScaleEstimateUpdate, CodeScaleEstimateResponse,
    ActivitySnapshotUpdate, ActivitySnapshotResponse,
    ReviewSnapshotUpdate, ReviewSnapshotResponse,
)
from app.services.estimator import generate_estimates

router = APIRouter(tags=["estimates"])


# ── Session ───────────────────────────────────────────────────────────────────

@router.post("/api/projects/{project_id}/sessions", status_code=201)
async def create_session(project_id: int):
    db = await get_db()
    async with db.execute("SELECT id FROM projects WHERE id=?", (project_id,)) as cur:
        if not await cur.fetchone():
            raise HTTPException(404, "Project not found")
    return await generate_estimates(db, project_id)


@router.get("/api/projects/{project_id}/sessions", response_model=list[SessionResponse])
async def list_sessions(project_id: int):
    db = await get_db()
    async with db.execute(
        "SELECT * FROM estimate_sessions WHERE project_id=? ORDER BY id DESC", (project_id,)
    ) as cur:
        rows = await cur.fetchall()
    return [dict(r) for r in rows]


@router.get("/api/sessions/{session_id}")
async def get_session(session_id: int):
    db = await get_db()
    async with db.execute("SELECT * FROM estimate_sessions WHERE id=?", (session_id,)) as cur:
        session = await cur.fetchone()
    if not session:
        raise HTTPException(404, "Session not found")

    async with db.execute(
        "SELECT * FROM sheet1_estimates WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        s1 = [dict(r) for r in await cur.fetchall()]

    async with db.execute(
        "SELECT * FROM code_scale_estimates WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        s2 = [dict(r) for r in await cur.fetchall()]

    async with db.execute(
        "SELECT * FROM activity_workload_snapshots WHERE session_id=? ORDER BY activity_type, id",
        (session_id,)
    ) as cur:
        acts = [dict(r) for r in await cur.fetchall()]

    async with db.execute(
        "SELECT * FROM review_workload_snapshots WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        revs = [dict(r) for r in await cur.fetchall()]

    return {
        "session": dict(session),
        "sheet1_estimates": s1,
        "code_scale_estimates": s2,
        "activity_snapshots": acts,
        "review_snapshots": revs,
    }


@router.delete("/api/sessions/{session_id}", status_code=204)
async def delete_session(session_id: int):
    db = await get_db()
    async with db.execute("SELECT id FROM estimate_sessions WHERE id=?", (session_id,)) as cur:
        if not await cur.fetchone():
            raise HTTPException(404, "Session not found")
    for table in (
        "sheet1_estimates",
        "code_scale_estimates",
        "activity_workload_snapshots",
        "review_workload_snapshots",
        "generated_files",
    ):
        await db.execute(f"DELETE FROM {table} WHERE session_id=?", (session_id,))
    await db.execute("DELETE FROM estimate_sessions WHERE id=?", (session_id,))
    await db.commit()


# ── Sheet 1 Estimates ─────────────────────────────────────────────────────────

@router.put("/api/sheet1-estimates/{row_id}", response_model=Sheet1EstimateResponse)
async def update_sheet1_estimate(row_id: int, data: Sheet1EstimateUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE sheet1_estimates SET {set_clause} WHERE id=?",
        list(fields.values()) + [row_id],
    )
    await db.commit()
    async with db.execute("SELECT * FROM sheet1_estimates WHERE id=?", (row_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Not found")
    return dict(row)


# ── Code Scale Estimates ──────────────────────────────────────────────────────

@router.put("/api/code-estimates/{row_id}", response_model=CodeScaleEstimateResponse)
async def update_code_estimate(row_id: int, data: CodeScaleEstimateUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE code_scale_estimates SET {set_clause} WHERE id=?",
        list(fields.values()) + [row_id],
    )
    await db.commit()
    async with db.execute("SELECT * FROM code_scale_estimates WHERE id=?", (row_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Not found")
    return dict(row)


# ── Activity Snapshots ────────────────────────────────────────────────────────

@router.put("/api/activity-snapshots/{row_id}", response_model=ActivitySnapshotResponse)
async def update_activity_snapshot(row_id: int, data: ActivitySnapshotUpdate):
    db = await get_db()
    if data.workload is None:
        raise HTTPException(400, "workload required")
    await db.execute(
        "UPDATE activity_workload_snapshots SET workload=? WHERE id=?",
        (data.workload, row_id),
    )
    await db.commit()
    async with db.execute(
        "SELECT * FROM activity_workload_snapshots WHERE id=?", (row_id,)
    ) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Not found")
    return dict(row)


# ── Review Snapshots ──────────────────────────────────────────────────────────

@router.put("/api/review-snapshots/{row_id}", response_model=ReviewSnapshotResponse)
async def update_review_snapshot(row_id: int, data: ReviewSnapshotUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE review_workload_snapshots SET {set_clause} WHERE id=?",
        list(fields.values()) + [row_id],
    )
    await db.commit()
    async with db.execute(
        "SELECT * FROM review_workload_snapshots WHERE id=?", (row_id,)
    ) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Not found")
    return dict(row)
