"""
Demo script showing example usage of the Project Scheduler
This creates the JIGAWA EXECUTION PLAN data exactly from the CSV file.
"""

import csv
from project_scheduler import Project, Activity, ActivitySection, CalendarFormat, ExcelGenerator, ProjectSchedulerGUI


def read_exact_csv_data():
    """Read the CSV file and extract activities exactly as they are"""
    activities = []
    
    with open('PROJECT Activities - Sch - Cost.xlsx - EXECUTION PLAN .csv', 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        rows = list(reader)
    
    # Find the actual data starting row (after headers)
    data_start_row = None
    for i, row in enumerate(rows):
        if len(row) > 0 and row[0] == 'S/N':
            data_start_row = i + 1
            break
    
    if data_start_row is None:
        print("Could not find data start row")
        return activities
    
    current_section = None
    
    for i, row in enumerate(rows[data_start_row:], start=data_start_row+1):
        if len(row) < 9:
            continue
            
        s_n = row[0].strip() if len(row) > 0 and row[0] else ""
        task = row[1].strip() if len(row) > 1 and row[1] else ""
        action = row[2].strip() if len(row) > 2 and row[2] else ""
        duration_str = row[3].strip() if len(row) > 3 and row[3] else ""
        precursor = row[4].strip() if len(row) > 4 and row[4] else ""
        sequence_str = row[5].strip() if len(row) > 5 and row[5] else ""
        schedule_str = row[6].strip() if len(row) > 6 and row[6] else ""
        resources = row[7].strip() if len(row) > 7 and row[7] else ""
        budget_str = row[8].strip() if len(row) > 8 and row[8] else ""
        
        # Determine section
        if "PRE-KICKOFF" in task.upper():
            current_section = ActivitySection.PRE_KICKOFF
            continue
        elif "POST KICKOFF" in task.upper() or task.upper() == "POST KICKOFF":
            current_section = ActivitySection.POST_KICKOFF
            continue
        elif any(keyword in task.upper() for keyword in ["PRE-INSTALLATION ACTIVITIES", "SHIPPING ACTIVITIES", 
                 "FINANCING ACTIVITIES", "INSTALLATION OF ALL IN ONE", "MINI-GRID INSTALLATION", 
                 "DEMOBILIZATION", "PROCUREMENT FOR", "PROGRESS- FINANCING ACTIVITIES", "SUB-T0TAL"]):
            # These are section headers, skip them
            continue
        
        # Skip if no meaningful task name, or if it's a number only, or empty
        if not task or s_n.isdigit() and not task.strip():
            continue
        if task.replace(" ", "").replace("-", "").replace(",", "") == "":
            continue
            
        # Parse duration (handle empty exactly as in CSV)
        try:
            duration = int(duration_str) if duration_str else 0
        except ValueError:
            duration = 0
            
        # Parse sequence (handle empty exactly as in CSV)
        try:
            sequence = int(sequence_str) if sequence_str else 0
        except ValueError:
            sequence = 0
            
        # Parse budget (handle empty and convert from millions exactly as in CSV)
        try:
            budget_millions = float(budget_str.replace(',', '')) if budget_str else 0.0
            budget = budget_millions * 1000000  # Convert to actual amount
        except ValueError:
            budget = 0.0
        
        # Default section if not determined
        if current_section is None:
            current_section = ActivitySection.PRE_KICKOFF if i < 20 else ActivitySection.POST_KICKOFF
        
        # Handle "Ditto" in resources (keep exactly as is)
        if resources.upper() == "DITTO":
            resources = "Ditto"
        
        activity = Activity(
            task=task,
            action_needed=action,  # Keep empty if empty in CSV
            duration=duration,
            precursor=precursor,   # Keep empty if empty in CSV
            sequence=sequence,
            resources=resources,   # Keep empty if empty in CSV
            budget=budget,
            section=current_section
        )
        
        activities.append(activity)
    
    return activities


def create_jigawa_project():
    """Create the JIGAWA EXECUTION PLAN project with exact CSV data"""
    
    # Create project
    project = Project(
        title="JIGAWA EXECUTION PLAN SCHEDULE",
        calendar_format=CalendarFormat.SIX_DAY
    )
    
    # Read activities exactly from CSV
    activities = read_exact_csv_data()
    
    # Add all activities to project exactly as they are in the CSV
    for activity in activities:
        project.add_activity(activity)
    
    return project



def main():
    """Generate demo Excel and text files with JIGAWA project data"""
    print("Creating JIGAWA EXECUTION PLAN project...")
    project = create_jigawa_project()
    
    print("Generating Excel and text files...")
    generator = ExcelGenerator(project)
    excel_path = "JIGAWA_execution_plan.xlsx"
    generator.generate(excel_path)
    
    # Generate accompanying .txt file (like GUI does)
    txt_path = "JIGAWA_execution_plan_details.txt"
    gui = ProjectSchedulerGUI()
    preview_text = gui.generate_preview_text(project)
    with open(txt_path, 'w', encoding='utf-8') as txt_file:
        txt_file.write(preview_text)
    
    print(f"JIGAWA project files created:")
    print(f"  • Excel: {excel_path}")
    print(f"  • Text:  {txt_path}")
    print("\nProject Summary:")
    print(f"Title: {project.title}")
    print(f"Calendar Format: {project.calendar_format.value}")
    print(f"Total Activities: {len(project.activities)}")
    
    pre_kickoff = project.get_activities_by_section(ActivitySection.PRE_KICKOFF)
    post_kickoff = project.get_activities_by_section(ActivitySection.POST_KICKOFF)
    
    print(f"Pre-Kickoff Activities: {len(pre_kickoff)}")
    print(f"Post Kick-off Activities: {len(post_kickoff)}")
    
    total_budget = sum(activity.budget for activity in project.activities)
    print(f"Total Budget: ₦{total_budget:,.2f} ({total_budget/1000000:.2f} million)")


if __name__ == "__main__":
    main() 