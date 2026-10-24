import aiosqlite
import os

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "estimatem.db")

_db: aiosqlite.Connection | None = None


async def get_db() -> aiosqlite.Connection:
    return _db


async def init_db():
    global _db
    _db = await aiosqlite.connect(DB_PATH)
    _db.row_factory = aiosqlite.Row
    await _db.execute("PRAGMA foreign_keys = ON")
    await _create_tables(_db)
    # 存量库迁移：补加 phase_wbs 列（新建库已包含，重复执行无害）
    try:
        await _db.execute("ALTER TABLE config_items ADD COLUMN phase_wbs TEXT")
        await _db.commit()
    except Exception:
        pass  # 列已存在，忽略
    await _db.commit()


async def close_db():
    global _db
    if _db:
        await _db.close()
        _db = None


async def _create_tables(db: aiosqlite.Connection):
    await db.executescript("""
CREATE TABLE IF NOT EXISTS projects (
  id           INTEGER PRIMARY KEY AUTOINCREMENT,
  name         TEXT NOT NULL,
  code         TEXT NOT NULL,
  leader       TEXT,
  cm_manager   TEXT,
  qa_manager   TEXT,
  round        INTEGER DEFAULT 1,
  estimators   TEXT,
  team_members TEXT,
  created_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
  updated_at   DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS config_items (
  id           INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id   INTEGER NOT NULL REFERENCES projects(id),
  seq          INTEGER,
  name         TEXT NOT NULL,
  identifier   TEXT,
  version      TEXT,
  item_type    TEXT,
  actual_pages INTEGER,
  workload     REAL,
  phase_wbs    TEXT
);

CREATE TABLE IF NOT EXISTS code_modules (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id  INTEGER NOT NULL REFERENCES projects(id),
  seq         INTEGER,
  name        TEXT,
  actual_loc  INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS wbs_tasks (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id    INTEGER NOT NULL REFERENCES projects(id),
  wbs           TEXT,
  name          TEXT NOT NULL,
  start_date    DATE,
  end_date      DATE,
  related_items TEXT,
  duration_days INTEGER,
  workload      REAL,
  resources     TEXT,
  task_type     TEXT
);

CREATE TABLE IF NOT EXISTS estimate_sessions (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id  INTEGER NOT NULL REFERENCES projects(id),
  created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS sheet1_estimates (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id      INTEGER NOT NULL REFERENCES estimate_sessions(id),
  config_item_id  INTEGER REFERENCES config_items(id),
  wbs_task_id     INTEGER REFERENCES wbs_tasks(id),
  item_name       TEXT NOT NULL,
  actual_pages    INTEGER,
  v1 INTEGER, v2 INTEGER, v3 INTEGER, v4 INTEGER,
  em_target       REAL,
  wl1 REAL, wl2 REAL, wl3 REAL, wl4 REAL
);

CREATE TABLE IF NOT EXISTS code_scale_estimates (
  id           INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id   INTEGER NOT NULL REFERENCES estimate_sessions(id),
  module_id    INTEGER REFERENCES code_modules(id),
  wbs_task_id  INTEGER REFERENCES wbs_tasks(id),
  module_name  TEXT NOT NULL,
  actual_loc   INTEGER,
  v1 INTEGER, v2 INTEGER, v3 INTEGER, v4 INTEGER,
  em_target    REAL,
  wl1 REAL, wl2 REAL, wl3 REAL, wl4 REAL
);

CREATE TABLE IF NOT EXISTS activity_workload_snapshots (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id    INTEGER NOT NULL REFERENCES estimate_sessions(id),
  task_id       INTEGER REFERENCES wbs_tasks(id),
  activity_type TEXT NOT NULL,
  phase_name    TEXT,
  activity_name TEXT NOT NULL,
  workload      REAL
);

CREATE TABLE IF NOT EXISTS review_workload_snapshots (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id      INTEGER NOT NULL REFERENCES estimate_sessions(id),
  task_id         INTEGER REFERENCES wbs_tasks(id),
  activity_name   TEXT NOT NULL,
  participants    INTEGER,
  total_hours     REAL
);

CREATE TABLE IF NOT EXISTS generated_files (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  project_id  INTEGER NOT NULL REFERENCES projects(id),
  session_id  INTEGER REFERENCES estimate_sessions(id),
  filename    TEXT,
  filepath    TEXT,
  created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
);
""")
