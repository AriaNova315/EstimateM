import os
from urllib.parse import quote
from fastapi import APIRouter, HTTPException
from fastapi.responses import Response

from app.database import get_db
from app.services.excel_builder import build_excel

router = APIRouter(tags=["export"])


@router.get("/api/sessions/{session_id}/export")
async def export_session(session_id: int):
    db = await get_db()
    async with db.execute(
        "SELECT es.id, p.name FROM estimate_sessions es JOIN projects p ON p.id=es.project_id WHERE es.id=?",
        (session_id,)
    ) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Session not found")

    session_id_val, project_name = row
    xlsx_bytes = await build_excel(db, session_id)

    safe_name = project_name.replace('/', '_').replace('\\', '_')
    filename = f"{safe_name}_规模及工作量估算表_session{session_id}.xlsx"

    encoded = quote(filename)
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"},
    )
