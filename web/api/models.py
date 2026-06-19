from pydantic import BaseModel
from typing import List, Optional
from datetime import date
from enum import Enum

# Using enums that match core_logic to ensure consistency
class CalendarFormatStr(str, Enum):
    FIVE_DAY = "5-day week"
    SIX_DAY = "6-day week"
    SEVEN_DAY = "7-day week"

class ActivitySectionStr(str, Enum):
    PRE_KICKOFF = "Pre-Kickoff Activities"
    POST_KICKOFF = "Post Kick-off Activities"

class ActivityCreate(BaseModel):
    task: str
    action_needed: str = ""
    duration: int = 0
    precursor: Optional[str] = None
    sequence: int = 0
    resources: str = ""
    budget: float = 0.0
    section: ActivitySectionStr

class ActivityUpdate(ActivityCreate):
    pass

class ActivityResponse(ActivityCreate):
    id: str
    project_id: str

class ProjectCreate(BaseModel):
    title: str
    start_date: Optional[date] = None
    calendar_format: CalendarFormatStr = CalendarFormatStr.FIVE_DAY
    logo_path: Optional[str] = None

class ProjectUpdate(ProjectCreate):
    pass

class ProjectResponse(ProjectCreate):
    id: str
    created_at: str
    updated_at: str
    activities: List[ActivityResponse] = []
