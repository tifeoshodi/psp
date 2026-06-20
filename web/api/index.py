from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
import uuid
import os
import tempfile
from typing import List

import os
import sys

# Ensure Vercel can find modules in the api directory
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from models import ProjectCreate, ProjectUpdate, ProjectResponse, ActivityCreate, ActivityUpdate, ActivityResponse
from database import get_db_connection
from core_logic import Project, Activity, ActivitySection, CalendarFormat, ExcelGenerator, GanttChartGenerator
import openpyxl

app = FastAPI(title="Project Scheduler API", docs_url="/api/docs", openapi_url="/api/openapi.json")

@app.get("/api/projects", response_model=List[ProjectResponse])
def list_projects():
    conn = get_db_connection()
    projects = conn.execute("SELECT * FROM projects").fetchall()
    result = []
    for p in projects:
        # Get activities for this project
        activities = conn.execute("SELECT * FROM activities WHERE project_id = ?", (p['id'],)).fetchall()
        p_dict = dict(p)
        p_dict['activities'] = [dict(a) for a in activities]
        result.append(p_dict)
    conn.close()
    return result

@app.post("/api/projects", response_model=ProjectResponse)
def create_project(project: ProjectCreate):
    conn = get_db_connection()
    project_id = str(uuid.uuid4())
    start_date_str = project.start_date.isoformat() if project.start_date else None
    
    conn.execute(
        "INSERT INTO projects (id, title, start_date, calendar_format, logo_path) VALUES (?, ?, ?, ?, ?)",
        (project_id, project.title, start_date_str, project.calendar_format.value, project.logo_path)
    )
    conn.commit()
    
    # Fetch created project
    p = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
    conn.close()
    
    p_dict = dict(p)
    p_dict['activities'] = []
    return p_dict

@app.get("/api/projects/{project_id}", response_model=ProjectResponse)
def get_project(project_id: str):
    conn = get_db_connection()
    p = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
    if not p:
        conn.close()
        raise HTTPException(status_code=404, detail="Project not found")
        
    activities = conn.execute("SELECT * FROM activities WHERE project_id = ?", (project_id,)).fetchall()
    conn.close()
    
    p_dict = dict(p)
    p_dict['activities'] = [dict(a) for a in activities]
    return p_dict

@app.put("/api/projects/{project_id}", response_model=ProjectResponse)
def update_project(project_id: str, project: ProjectUpdate):
    conn = get_db_connection()
    start_date_str = project.start_date.isoformat() if project.start_date else None
    
    conn.execute(
        "UPDATE projects SET title = ?, start_date = ?, calendar_format = ?, logo_path = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (project.title, start_date_str, project.calendar_format.value, project.logo_path, project_id)
    )
    conn.commit()
    
    p = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
    conn.close()
    if not p:
        raise HTTPException(status_code=404, detail="Project not found")
        
    p_dict = dict(p)
    p_dict['activities'] = []
    return p_dict

@app.delete("/api/projects/{project_id}")
def delete_project(project_id: str):
    conn = get_db_connection()
    conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/api/projects/{project_id}/activities", response_model=ActivityResponse)
def add_activity(project_id: str, activity: ActivityCreate):
    conn = get_db_connection()
    activity_id = str(uuid.uuid4())
    
    conn.execute(
        """INSERT INTO activities 
        (id, project_id, task, action_needed, duration, precursor, sequence, resources, budget, section) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (activity_id, project_id, activity.task, activity.action_needed, activity.duration,
         activity.precursor, activity.sequence, activity.resources, activity.budget, activity.section.value)
    )
    conn.commit()
    
    a = conn.execute("SELECT * FROM activities WHERE id = ?", (activity_id,)).fetchone()
    conn.close()
    return dict(a)

@app.delete("/api/projects/{project_id}/activities/{activity_id}")
def delete_activity(project_id: str, activity_id: str):
    conn = get_db_connection()
    conn.execute("DELETE FROM activities WHERE id = ? AND project_id = ?", (activity_id, project_id))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/api/projects/{project_id}/generate-excel")
def generate_excel(project_id: str, background_tasks: BackgroundTasks):
    conn = get_db_connection()
    p = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
    if not p:
        conn.close()
        raise HTTPException(status_code=404, detail="Project not found")
        
    activities_db = conn.execute("SELECT * FROM activities WHERE project_id = ?", (project_id,)).fetchall()
    conn.close()
    
    # Reconstruct core_logic Project
    import datetime
    start_date = None
    if p['start_date']:
        start_date = datetime.date.fromisoformat(p['start_date'])
        
    # Map calendar format
    cal_format = CalendarFormat.FIVE_DAY
    if p['calendar_format'] == '6-day week':
        cal_format = CalendarFormat.SIX_DAY
    elif p['calendar_format'] == '7-day week':
        cal_format = CalendarFormat.SEVEN_DAY
        
    core_project = Project(
        title=p['title'],
        calendar_format=cal_format,
        start_date=start_date
    )
    
    for a in activities_db:
        section = ActivitySection.PRE_KICKOFF if a['section'] == 'Pre-Kickoff Activities' else ActivitySection.POST_KICKOFF
        activity = Activity(
            task=a['task'],
            action_needed=a['action_needed'] or "",
            duration=a['duration'],
            precursor=a['precursor'] or "",
            sequence=a['sequence'],
            resources=a['resources'] or "",
            budget=a['budget'],
            section=section
        )
        core_project.add_activity(activity)
        
    # Generate Excel
    generator = ExcelGenerator(core_project, custom_logo_path=p['logo_path'])
    
    # Save to a temporary file
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    output_path = temp_file.name
    temp_file.close()
    
    generator.generate(output_path)
    
    # Return file and schedule cleanup
    def cleanup_file(path: str):
        if os.path.exists(path):
            os.remove(path)
            
    background_tasks.add_task(cleanup_file, output_path)
    
    filename = f"{core_project.title.replace(' ', '_')}_Schedule.xlsx"
    return FileResponse(path=output_path, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.get("/api/projects/{project_id}/gantt")
def get_gantt_data(project_id: str):
    conn = get_db_connection()
    p = conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()
    if not p:
        conn.close()
        raise HTTPException(status_code=404, detail="Project not found")
        
    activities_db = conn.execute("SELECT * FROM activities WHERE project_id = ?", (project_id,)).fetchall()
    conn.close()
    
    # Reconstruct core_logic Project
    import datetime
    start_date = None
    if p['start_date']:
        start_date = datetime.date.fromisoformat(p['start_date'])
    else:
        start_date = datetime.date.today()
        
    # Map calendar format
    cal_format = CalendarFormat.FIVE_DAY
    if p['calendar_format'] == '6-day week':
        cal_format = CalendarFormat.SIX_DAY
    elif p['calendar_format'] == '7-day week':
        cal_format = CalendarFormat.SEVEN_DAY
        
    core_project = Project(
        title=p['title'],
        calendar_format=cal_format,
        start_date=start_date
    )
    
    for a in activities_db:
        section = ActivitySection.PRE_KICKOFF if 'Pre' in a['section'] else ActivitySection.POST_KICKOFF
        activity = Activity(
            task=a['task'],
            action_needed=a['action_needed'] or "",
            duration=a['duration'],
            precursor=a['precursor'] or "",
            sequence=a['sequence'],
            resources=a['resources'] or "",
            budget=a['budget'],
            section=section
        )
        core_project.add_activity(activity)
        
    # Generate timeline data
    # We use a dummy workbook just to instantiate the generator
    dummy_wb = openpyxl.Workbook()
    generator = GanttChartGenerator(core_project, dummy_wb)
    timeline_data = generator._calculate_timeline_data(start_date)
    
    # Format the data for JSON response
    response_data = {
        'project_start_date': timeline_data['project_start_date'].isoformat(),
        'timeline_start': timeline_data['timeline_start'].isoformat(),
        'timeline_days': timeline_data['timeline_days'],
        'date_timeline': [d.isoformat() for d in timeline_data['date_timeline']],
        'tasks': []
    }
    
    for task in timeline_data['task_timeline']:
        response_data['tasks'].append({
            'id': str(id(task['activity'])),
            'name': task['activity'].task,
            'start_date': task['start_date'].isoformat(),
            'end_date': task['end_date'].isoformat(),
            'start_day': task['start_day'],
            'end_day': task['end_day'],
            'duration': task['duration'],
            'task_type': task['task_type'],
            'is_critical': task['is_critical']
        })
        
    return response_data
