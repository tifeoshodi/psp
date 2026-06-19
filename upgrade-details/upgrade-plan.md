Phase 1: The UI/UX Design Sprint (Planning the Screens)
To translate the current single-window tkinter app into a proper, multi-view internal tool, we need to break the user journey into distinct, modular screens.

Here is the proposed screen architecture for Stitch to generate:

1. The Global Dashboard (New)

Purpose: The landing page to view, load, and manage all saved projects from the SQLite database.

UI Elements: A data table listing projects (Name, Start Date, Total Activities, Total Budget, Last Modified). Actions to "Create New Project", "Edit", "Delete", or "Duplicate".

2. The Project Workspace (The Core Editor)

Purpose: Replaces the current tkinter input form and treeview. This is where data entry happens.

UI Elements: * Top Bar: Project Title, Start Date picker, Calendar Format toggle, and a "Save to DB" button.

Left Sidebar (or Top Form): The form to add a new Activity (Task, Action, Duration, Precursor, Sequence, Resources, Budget, Section).

Main Pane: A robust, interactive data table (using the horizontal scrolling and sticky headers defined in our Antigravity rules) displaying the activities. It needs inline editing or a slide-out drawer for modifying existing activities.

3. The Preview & Export Modal/Screen

Purpose: Replaces the pop-up text preview and handles the Excel generation.

UI Elements: A split view. The left side shows the generated text summary (Budget, Critical Path, etc.). The right side contains the "Custom Logo Upload" component, the final "Generate Excel" button, and a visual loading skeleton/spinner to handle the backend processing time.

Phase 2: Database Architecture (SQLite Schema)
We will map your current Python dataclasses directly to relational tables.

Table 1: projects

id (PK, UUID or Integer)

title (Text)

start_date (Date)

calendar_format (Text - "5-day", "6-day", "7-day")

logo_path (Text, nullable)

created_at / updated_at (Timestamps)

Table 2: activities

id (PK, UUID or Integer)

project_id (FK -> projects.id)

task (Text)

action_needed (Text)

duration (Integer)

precursor (Text)

sequence (Integer)

resources (Text)

budget (Float)

section (Text - "Pre-Kickoff", "Post Kick-off")

Phase 3: Backend Migration
Strip out all GUI code (tkinter imports, ProjectSchedulerGUI class) from project_scheduler.py.

Isolate the Project, Activity, ScheduleCalculator, ExcelGenerator, and GanttChartGenerator classes into a core_logic.py module.

Write FastAPI routes (POST /projects, GET /projects, POST /projects/{id}/generate-excel) that interact with the SQLite DB and pass the data into your existing ExcelGenerator.

Phase 4: Frontend Scaffolding via Stitch MCP
We will feed the screen definitions from Phase 1 into the Stitch MCP, strictly enforcing the desktop-first, modular components, strict theming, a11y focus, as it is spelt out in the global workflows.