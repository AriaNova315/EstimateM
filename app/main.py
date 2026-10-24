from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from contextlib import asynccontextmanager
import os

from app.database import init_db, close_db
from app.routers import projects, config_items, wbs_tasks, estimates, export


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    yield
    await close_db()


app = FastAPI(title="EstimateM", lifespan=lifespan)

static_dir = os.path.join(os.path.dirname(__file__), "static")
app.mount("/static", StaticFiles(directory=static_dir), name="static")

app.include_router(projects.router)
app.include_router(config_items.router)
app.include_router(wbs_tasks.router)
app.include_router(estimates.router)
app.include_router(export.router)


@app.get("/")
async def index():
    return FileResponse(os.path.join(static_dir, "index.html"))


@app.get("/health")
async def health():
    return {"status": "ok"}
