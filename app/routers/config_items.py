from fastapi import APIRouter, HTTPException

from app.database import get_db
from app.models import ConfigItemUpdate, ConfigItemResponse, CodeModuleUpdate, CodeModuleResponse

router = APIRouter(tags=["config_items"])


@router.get("/api/projects/{project_id}/config-items", response_model=list[ConfigItemResponse])
async def list_config_items(project_id: int):
    db = await get_db()
    async with db.execute(
        "SELECT * FROM config_items WHERE project_id=? ORDER BY seq, id", (project_id,)
    ) as cur:
        rows = await cur.fetchall()
    return [dict(r) for r in rows]


@router.put("/api/config-items/{item_id}", response_model=ConfigItemResponse)
async def update_config_item(item_id: int, data: ConfigItemUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE config_items SET {set_clause} WHERE id=?",
        list(fields.values()) + [item_id],
    )
    await db.commit()
    async with db.execute("SELECT * FROM config_items WHERE id=?", (item_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Config item not found")
    return dict(row)


@router.get("/api/projects/{project_id}/code-modules", response_model=list[CodeModuleResponse])
async def list_code_modules(project_id: int):
    db = await get_db()
    async with db.execute(
        "SELECT * FROM code_modules WHERE project_id=? ORDER BY seq, id", (project_id,)
    ) as cur:
        rows = await cur.fetchall()
    return [dict(r) for r in rows]


@router.put("/api/code-modules/{module_id}", response_model=CodeModuleResponse)
async def update_code_module(module_id: int, data: CodeModuleUpdate):
    db = await get_db()
    fields = {k: v for k, v in data.model_dump().items() if v is not None}
    if not fields:
        raise HTTPException(400, "No fields to update")
    set_clause = ", ".join(f"{k}=?" for k in fields)
    await db.execute(
        f"UPDATE code_modules SET {set_clause} WHERE id=?",
        list(fields.values()) + [module_id],
    )
    await db.commit()
    async with db.execute("SELECT * FROM code_modules WHERE id=?", (module_id,)) as cur:
        row = await cur.fetchone()
    if not row:
        raise HTTPException(404, "Code module not found")
    return dict(row)
