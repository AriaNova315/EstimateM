from pydantic import BaseModel
from typing import Optional


class ProjectCreate(BaseModel):
    name: str
    code: str
    leader: Optional[str] = None
    cm_manager: Optional[str] = None
    qa_manager: Optional[str] = None
    round: int = 1
    estimators: Optional[str] = None
    team_members: Optional[str] = None


class ProjectUpdate(BaseModel):
    name: Optional[str] = None
    code: Optional[str] = None
    leader: Optional[str] = None
    cm_manager: Optional[str] = None
    qa_manager: Optional[str] = None
    round: Optional[int] = None
    estimators: Optional[str] = None
    team_members: Optional[str] = None


class ProjectResponse(BaseModel):
    id: int
    name: str
    code: str
    leader: Optional[str]
    cm_manager: Optional[str]
    qa_manager: Optional[str]
    round: int
    estimators: Optional[str]
    team_members: Optional[str]
    created_at: str
    updated_at: str


class ConfigItemUpdate(BaseModel):
    item_type: Optional[str] = None
    actual_pages: Optional[int] = None
    workload: Optional[float] = None


class ConfigItemResponse(BaseModel):
    id: int
    project_id: int
    seq: Optional[int]
    name: str
    identifier: Optional[str]
    version: Optional[str]
    item_type: Optional[str]
    actual_pages: Optional[int]
    workload: Optional[float]


class CodeModuleUpdate(BaseModel):
    name: Optional[str] = None
    actual_loc: Optional[int] = None


class CodeModuleResponse(BaseModel):
    id: int
    project_id: int
    seq: Optional[int]
    name: Optional[str]
    actual_loc: int


class WbsTaskUpdate(BaseModel):
    workload: Optional[float] = None
    task_type: Optional[str] = None


class WbsTaskResponse(BaseModel):
    id: int
    project_id: int
    wbs: Optional[str]
    name: str
    start_date: Optional[str]
    end_date: Optional[str]
    related_items: Optional[str]
    duration_days: Optional[int]
    workload: Optional[float]
    resources: Optional[str]
    task_type: Optional[str]
    latest_est_workload: Optional[float] = None


# ── Estimate Session ──────────────────────────────────────────────────────────

class SessionResponse(BaseModel):
    id: int
    project_id: int
    created_at: str


class Sheet1EstimateUpdate(BaseModel):
    v1: Optional[int] = None
    v2: Optional[int] = None
    v3: Optional[int] = None
    v4: Optional[int] = None
    em_target: Optional[float] = None
    wl1: Optional[float] = None
    wl2: Optional[float] = None
    wl3: Optional[float] = None
    wl4: Optional[float] = None


class Sheet1EstimateResponse(BaseModel):
    id: int
    session_id: int
    config_item_id: Optional[int]
    wbs_task_id: Optional[int]
    item_name: str
    actual_pages: Optional[int]
    v1: Optional[int]
    v2: Optional[int]
    v3: Optional[int]
    v4: Optional[int]
    em_target: Optional[float]
    wl1: Optional[float]
    wl2: Optional[float]
    wl3: Optional[float]
    wl4: Optional[float]


class CodeScaleEstimateUpdate(BaseModel):
    v1: Optional[int] = None
    v2: Optional[int] = None
    v3: Optional[int] = None
    v4: Optional[int] = None
    em_target: Optional[float] = None
    wl1: Optional[float] = None
    wl2: Optional[float] = None
    wl3: Optional[float] = None
    wl4: Optional[float] = None


class CodeScaleEstimateResponse(BaseModel):
    id: int
    session_id: int
    module_id: Optional[int]
    wbs_task_id: Optional[int]
    module_name: str
    actual_loc: Optional[int]
    v1: Optional[int]
    v2: Optional[int]
    v3: Optional[int]
    v4: Optional[int]
    em_target: Optional[float]
    wl1: Optional[float]
    wl2: Optional[float]
    wl3: Optional[float]
    wl4: Optional[float]


class ActivitySnapshotUpdate(BaseModel):
    workload: Optional[float] = None


class ActivitySnapshotResponse(BaseModel):
    id: int
    session_id: int
    task_id: Optional[int]
    activity_type: str
    phase_name: Optional[str]
    activity_name: str
    workload: Optional[float]


class ReviewSnapshotUpdate(BaseModel):
    participants: Optional[int] = None
    total_hours: Optional[float] = None


class ReviewSnapshotResponse(BaseModel):
    id: int
    session_id: int
    task_id: Optional[int]
    activity_name: str
    participants: Optional[int]
    total_hours: Optional[float]
