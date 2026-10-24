from fastapi import APIRouter, HTTPException, UploadFile, File

from app.database import get_db
from app.models import ProjectCreate, ProjectUpdate, ProjectResponse
from app.services.importer import import_excel, peek_project_code

router = APIRouter(prefix="/api/projects", tags=["projects"])


@router.get("", response_model=list[ProjectResponse])
async def list_projects():
    db = await get_db()
    async with db.execute("SELECT * FROM projects ORDER BY id") as cur:
        rows = await cur.fetchall()
    return [dict(r) for r in rows]


@router.post("", response_model=ProjectResponse, status_code=201)
async def create_project(data: ProjectCreate):
    db = await get_db()
    async with db.execute(
        """INSERT INTO projects (name, code, leader, cm_manager, qa_manager, round, estimators, team_members)
           VALUES (?,?,?,?,?,?,?,?)""",
        (data.name, data.code, data.leader, data.cm_manager, data.qa_manager,
         data.round, data.estimators, data.team_members),
    ) as cur:
        project_id = cur.lastrowid
    await db.commit()
    async with db.execute("SELECT * FROM projects WHERE id=?", (project_id,)) as cur:
        row = await cur.fetchone()
    return dict(row)


@router.get("/{project_id}", response_model=ProjectResponse)
async def get_project(project_id: int):
    db = await get_db()
    async with db.execute("SELECT * FROM projects WHERE id=?", (project_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Project not found")
    return dict(row)


@router.put("/{project_id}", response_model=ProjectResponse)
async def update_project(project_id: int, data: ProjectUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE projects SET {set_clause}, updated_at=CURRENT_TIMESTAMP WHERE id=?",
        list(fields.values()) + [project_id],
    )
    await db.commit()
    async with db.execute("SELECT * FROM projects WHERE id=?", (project_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Project not found")
    return dict(row)


@router.delete("/{project_id}", status_code=204)
async def delete_project(project_id: int):
    db = await get_db()
    await db.execute("DELETE FROM projects WHERE id=?", (project_id,))
    await db.commit()


@router.post("/{project_id}/import")
async def import_project_excel(project_id: int, file: UploadFile = File(...)):
    db = await get_db()
    async with db.execute("SELECT id FROM projects WHERE id=?", (project_id,)) as cur:
        if not await cur.fetchone():
            raise HTTPException(404, "Project not found")
    content = await file.read()
    return await import_excel(db, project_id, content)


@router.post("/import-smart")
async def import_smart(file: UploadFile = File(...)):
    """从 Excel 导入：按 code 去重，已有则更新，无则新建。"""
    db = await get_db()
    content = await file.read()

    code = peek_project_code(content)
    project_id = None
    created = False

    if code:
        async with db.execute("SELECT id FROM projects WHERE code=?", (code,)) as cur:
            row = await cur.fetchone()
        if row:
            project_id = row["id"]

    if project_id is None:
        async with db.execute(
            """INSERT INTO projects (name, code) VALUES (?,?)""",
            ("导入中…", code or "IMPORT"),
        ) as cur:
            project_id = cur.lastrowid
        await db.commit()
        created = True

    result = await import_excel(db, project_id, content)
    result["project_id"] = project_id
    result["created"] = created
    return result
