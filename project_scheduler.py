"""
Project Scheduler Application
Creates Excel schedules from user input with advanced formatting and calculations.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple
from enum import Enum
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class CalendarFormat(Enum):
    FIVE_DAY = "5-day week"
    SIX_DAY = "6-day week"
    SEVEN_DAY = "7-day week"


class ActivitySection(Enum):
    PRE_KICKOFF = "Pre-Kickoff Activities"
    POST_KICKOFF = "Post Kick-off Activities"


@dataclass
class Activity:
    """Represents a single project activity/task"""
    task: str
    action_needed: str
    duration: int
    precursor: str
    sequence: int
    resources: str
    budget: float
    section: ActivitySection

    def __post_init__(self):
        # Ensure duration is positive
        if self.duration < 0:
            self.duration = 0


@dataclass
class Project:
    """Represents the entire project with all activities"""
    title: str
    activities: List[Activity] = field(default_factory=list)
    calendar_format: CalendarFormat = CalendarFormat.FIVE_DAY

    def add_activity(self, activity: Activity):
        """Add an activity to the project"""
        self.activities.append(activity)

    def get_activities_by_section(self, section: ActivitySection) -> List[Activity]:
        """Get all activities in a specific section"""
        return [activity for activity in self.activities if activity.section == section]

    def get_sequences_by_section(self, section: ActivitySection) -> List[int]:
        """Get unique sequences in a section, sorted"""
        activities = self.get_activities_by_section(section)
        sequences = list(set(activity.sequence for activity in activities))
        return sorted(sequences)


class ScheduleCalculator:
    """Handles schedule calculations and calendar adjustments"""

    @staticmethod
    def calculate_schedules(project: Project) -> Dict[int, int]:
        """Calculate schedule values for each sequence"""
        schedules = {}

        # Pre-kickoff activities always have schedule = 0
        pre_kickoff_sequences = project.get_sequences_by_section(ActivitySection.PRE_KICKOFF)
        for seq in pre_kickoff_sequences:
            schedules[seq] = 0

        # Post-kickoff activities have incremental schedules
        post_kickoff_sequences = project.get_sequences_by_section(ActivitySection.POST_KICKOFF)
        post_kickoff_activities = project.get_activities_by_section(ActivitySection.POST_KICKOFF)

        cumulative_schedule = 0
        for seq in post_kickoff_sequences:
            # Find max duration in this sequence
            seq_activities = [a for a in post_kickoff_activities if a.sequence == seq]
            max_duration = max(a.duration for a in seq_activities) if seq_activities else 0

            # Add max duration to cumulative schedule
            cumulative_schedule += max_duration
            schedules[seq] = cumulative_schedule

        return schedules

    @staticmethod
    def apply_calendar_format(schedule_days: int, calendar_format: CalendarFormat) -> int:
        """Apply calendar format adjustments to schedule days"""
        if calendar_format == CalendarFormat.FIVE_DAY:
            # Add 2 days per week for weekends
            weeks = schedule_days // 5
            extra_days = weeks * 2
            return schedule_days + extra_days
        elif calendar_format == CalendarFormat.SIX_DAY:
            # Add 1 day per week for single non-working day
            weeks = schedule_days // 6
            extra_days = weeks * 1
            return schedule_days + extra_days
        else:  # SEVEN_DAY
            return schedule_days

    @staticmethod
    def get_max_duration_activities(project: Project) -> List[Tuple[Activity, int]]:
        """Get activities that have max duration within their sequence group"""
        max_duration_activities = []

        for section in [ActivitySection.PRE_KICKOFF, ActivitySection.POST_KICKOFF]:
            sequences = project.get_sequences_by_section(section)
            section_activities = project.get_activities_by_section(section)

            for seq in sequences:
                seq_activities = [a for a in section_activities if a.sequence == seq]
                if seq_activities:
                    max_duration = max(a.duration for a in seq_activities)
                    for activity in seq_activities:
                        if activity.duration == max_duration:
                            max_duration_activities.append((activity, max_duration))

        return max_duration_activities


class ExcelGenerator:
    """Generates Excel files with proper formatting and calculations"""

    def __init__(self, project: Project):
        self.project = project
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "Project Schedule"

        # Styling constants
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.header_font = Font(color="FFFFFF", bold=True)
        self.section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        self.section_font = Font(bold=True)
        self.red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    def generate(self, output_path: str):
        """Generate the complete Excel file"""
        current_row = 1

        # Add project title
        current_row = self._add_project_title(current_row)
        current_row += 1  # Empty row

        # Add column headers
        current_row = self._add_headers(current_row)

        # Add Pre-Kickoff Activities section
        current_row = self._add_section(ActivitySection.PRE_KICKOFF, current_row)
        current_row += 1  # Empty row

        # Add Post Kick-off Activities section
        current_row = self._add_section(ActivitySection.POST_KICKOFF, current_row)
        
        # Add budget total row
        current_row = self._add_budget_total(current_row)

        # Apply formatting
        self._apply_formatting()

        # Save file
        self.workbook.save(output_path)

    def _add_project_title(self, start_row: int) -> int:
        """Add project title to the worksheet"""
        cell = self.worksheet.cell(row=start_row, column=1, value=f"Project Title: {self.project.title}")
        cell.font = Font(size=14, bold=True)
        self.worksheet.merge_cells(f"A{start_row}:H{start_row}")
        return start_row + 1

    def _add_headers(self, start_row: int) -> int:
        """Add column headers"""
        headers = [
            "Activities/Tasks", "Action Needed", "Duration", "Precursor",
            "Sequence", "Schedule (in days)", "Resources", "Budget"
        ]

        for col, header in enumerate(headers, 1):
            cell = self.worksheet.cell(row=start_row, column=col, value=header)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        return start_row + 1

    def _add_section(self, section: ActivitySection, start_row: int) -> int:
        """Add a section with its activities"""
        # Add section header
        section_cell = self.worksheet.cell(row=start_row, column=1, value=section.value)
        section_cell.fill = self.section_fill
        section_cell.font = self.section_font
        self.worksheet.merge_cells(f"A{start_row}:H{start_row}")
        current_row = start_row + 1

        # Get activities and schedules for this section
        activities = self.project.get_activities_by_section(section)
        schedules = ScheduleCalculator.calculate_schedules(self.project)
        max_duration_activities = ScheduleCalculator.get_max_duration_activities(self.project)
        max_duration_set = {id(activity) for activity, _ in max_duration_activities}

        # Sort activities by sequence
        activities.sort(key=lambda x: x.sequence)

        # Add activities
        for activity in activities:
            self._add_activity_row(activity, current_row, schedules, max_duration_set)
            current_row += 1

        # Merge cells for same sequences
        self._merge_schedule_cells(section, start_row + 1, current_row - 1)

        return current_row

    def _add_activity_row(self, activity: Activity, row: int, schedules: Dict[int, int], max_duration_set: set):
        """Add a single activity row"""
        # Calculate adjusted schedule - Pre-Kickoff activities always have 0
        if activity.section == ActivitySection.PRE_KICKOFF:
            adjusted_schedule = 0
        else:
            raw_schedule = schedules.get(activity.sequence, 0)
            adjusted_schedule = ScheduleCalculator.apply_calendar_format(raw_schedule, self.project.calendar_format)

        values = [
            activity.task,
            activity.action_needed,
            activity.duration,
            activity.precursor,
            activity.sequence,
            adjusted_schedule,
            activity.resources,
            activity.budget
        ]

        for col, value in enumerate(values, 1):
            cell = self.worksheet.cell(row=row, column=col, value=value)

            # Highlight max duration text in red
            if col == 3 and id(activity) in max_duration_set:  # Duration column
                cell.font = Font(color="FF0000", bold=True)  # Red text, bold

    def _merge_schedule_cells(self, section: ActivitySection, start_row: int, end_row: int):
        """Merge cells in Schedule column for same sequence values"""
        activities = self.project.get_activities_by_section(section)
        sequences = self.project.get_sequences_by_section(section)

        for seq in sequences:
            seq_activities = [a for a in activities if a.sequence == seq]
            if len(seq_activities) > 1:
                # Find the rows for this sequence
                seq_rows = []
                current_activities = activities[:]
                current_activities.sort(key=lambda x: x.sequence)

                row_counter = start_row
                for activity in current_activities:
                    if activity.sequence == seq:
                        seq_rows.append(row_counter)
                    row_counter += 1

                if len(seq_rows) > 1:
                    first_row = min(seq_rows)
                    last_row = max(seq_rows)
                    self.worksheet.merge_cells(f"F{first_row}:F{last_row}")

    def _add_budget_total(self, start_row: int) -> int:
        """Add a total row for budget calculation"""
        current_row = start_row + 1  # Add empty row before total
        
        # Add "Total" label
        total_cell = self.worksheet.cell(row=current_row, column=7, value="Total:")
        total_cell.font = Font(bold=True)
        total_cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Calculate total budget
        total_budget = sum(activity.budget for activity in self.project.activities)
        budget_cell = self.worksheet.cell(row=current_row, column=8, value=total_budget)
        budget_cell.font = Font(bold=True)
        budget_cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        budget_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        return current_row + 1

    def _apply_formatting(self):
        """Apply general formatting to the worksheet"""
        # Auto-adjust column widths
        for column in self.worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            adjusted_width = min(max_length + 2, 50)
            self.worksheet.column_dimensions[column_letter].width = adjusted_width

        # Apply borders to all used cells
        for row in self.worksheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.border = self.border
                    if not cell.fill.start_color.rgb and cell.row > 2:  # Don't override existing fills
                        cell.alignment = Alignment(horizontal='center', vertical='center')


class ProjectSchedulerGUI:
    """Main GUI application for project scheduling"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Project Scheduler")
        self.root.geometry("1200x800")  # Increased size to accommodate all elements
        self.root.minsize(1000, 700)  # Set minimum size
        
        self.project = None
        self.activities_data = []

        self.setup_ui()

    def setup_ui(self):
        """Set up the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)  # Make activities section expandable

        # Project title section
        self.setup_project_section(main_frame)

        # Calendar format section
        self.setup_calendar_section(main_frame)

        # Activities section
        self.setup_activities_section(main_frame)

        # Buttons section
        self.setup_buttons_section(main_frame)

    def setup_project_section(self, parent):
        """Set up project title input section"""
        # Project Title
        ttk.Label(parent, text="Project Title:", font=('Arial', 12, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10))

        self.project_title_var = tk.StringVar()
        project_entry = ttk.Entry(parent, textvariable=self.project_title_var, width=50, font=('Arial', 11))
        project_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

    def setup_calendar_section(self, parent):
        """Set up calendar format selection"""
        ttk.Label(parent, text="Calendar Format:", font=('Arial', 12, 'bold')).grid(
            row=1, column=0, sticky=tk.W, pady=(0, 10))

        self.calendar_format_var = tk.StringVar(value=CalendarFormat.FIVE_DAY.value)
        calendar_frame = ttk.Frame(parent)
        calendar_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 10))

        for i, fmt in enumerate(CalendarFormat):
            ttk.Radiobutton(calendar_frame, text=fmt.value, variable=self.calendar_format_var,
                          value=fmt.value).grid(row=0, column=i, padx=(0, 20), sticky=tk.W)

    def setup_activities_section(self, parent):
        """Set up activities input section"""
        # Activities section label
        ttk.Label(parent, text="Activities:", font=('Arial', 12, 'bold')).grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=(20, 10))

        # Activities frame
        activities_frame = ttk.Frame(parent)
        activities_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        activities_frame.columnconfigure(0, weight=1)
        activities_frame.rowconfigure(1, weight=1)

        # Activities input form
        self.setup_activity_form(activities_frame)

        # Activities list
        self.setup_activities_list(activities_frame)

    def setup_activity_form(self, parent):
        """Set up the form for adding new activities"""
        form_frame = ttk.LabelFrame(parent, text="Add New Activity", padding="10")
        form_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        form_frame.columnconfigure(1, weight=1)
        form_frame.columnconfigure(3, weight=1)

        # Form fields
        fields = [
            ("Task:", "task_var"),
            ("Action Needed:", "action_var"),
            ("Duration (days):", "duration_var"),
            ("Precursor:", "precursor_var"),
            ("Sequence:", "sequence_var"),
            ("Resources:", "resources_var"),
            ("Budget:", "budget_var")
        ]

        self.form_vars = {}
        for i, (label, var_name) in enumerate(fields):
            row = i // 2
            col = (i % 2) * 2

            ttk.Label(form_frame, text=label).grid(row=row, column=col, sticky=tk.W, padx=(0, 5), pady=2)

            if var_name in ['duration_var', 'sequence_var', 'budget_var']:
                self.form_vars[var_name] = tk.StringVar()
                entry = ttk.Entry(form_frame, textvariable=self.form_vars[var_name], width=20)
            else:
                self.form_vars[var_name] = tk.StringVar()
                entry = ttk.Entry(form_frame, textvariable=self.form_vars[var_name], width=30)

            entry.grid(row=row, column=col+1, sticky=(tk.W, tk.E), padx=(0, 20), pady=2)

        # Section selection
        ttk.Label(form_frame, text="Section:").grid(row=4, column=0, sticky=tk.W, padx=(0, 5), pady=2)
        self.section_var = tk.StringVar(value=ActivitySection.PRE_KICKOFF.value)
        section_combo = ttk.Combobox(form_frame, textvariable=self.section_var,
                                   values=[s.value for s in ActivitySection], state="readonly", width=25)
        section_combo.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=(0, 20), pady=2)

        # Add button
        ttk.Button(form_frame, text="Add Activity", command=self.add_activity).grid(
            row=4, column=3, padx=(20, 0), pady=10)

    def setup_activities_list(self, parent):
        """Set up the list view for activities"""
        self.list_frame = ttk.LabelFrame(parent, text="Activities List (0 activities)", padding="10")
        self.list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.list_frame.columnconfigure(0, weight=1)
        self.list_frame.rowconfigure(0, weight=1)

        # Treeview for activities
        columns = ("Section", "Task", "Action", "Duration", "Precursor", "Sequence", "Resources", "Budget")
        self.activities_tree = ttk.Treeview(self.list_frame, columns=columns, show="headings", height=15)

        # Configure columns
        for col in columns:
            self.activities_tree.heading(col, text=col)
            self.activities_tree.column(col, width=100, minwidth=80)

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.activities_tree.yview)
        h_scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.HORIZONTAL, command=self.activities_tree.xview)
        self.activities_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid layout
        self.activities_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Delete button and right-click context menu
        delete_frame = ttk.Frame(self.list_frame)
        delete_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W, tk.E))
        
        ttk.Button(delete_frame, text="Delete Selected", command=self.delete_activity_with_confirmation).pack(
            side=tk.LEFT, padx=(0, 10))
        
        # Setup right-click context menu
        self.setup_context_menu()

    def setup_buttons_section(self, parent):
        """Set up action buttons"""
        buttons_frame = ttk.LabelFrame(parent, text="Actions", padding="10")
        buttons_frame.grid(row=4, column=0, columnspan=2, pady=(20, 10), sticky=(tk.W, tk.E))

        # Create a centered button layout
        button_container = ttk.Frame(buttons_frame)
        button_container.pack(expand=True)

        ttk.Button(button_container, text="Preview Schedule", command=self.preview_schedule).pack(
            side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_container, text="Generate Excel", command=self.generate_excel, 
                  style="Accent.TButton" if hasattr(ttk.Style(), 'theme_use') else None).pack(
            side=tk.LEFT, padx=(0, 15))
        ttk.Button(button_container, text="Clear All", command=self.clear_all_with_confirmation).pack(
            side=tk.LEFT, padx=(0, 0))

    def add_activity(self):
        """Add a new activity to the list"""
        try:
            # Validate inputs
            task = self.form_vars['task_var'].get().strip()
            action = self.form_vars['action_var'].get().strip()
            duration_str = self.form_vars['duration_var'].get().strip()
            precursor = self.form_vars['precursor_var'].get().strip()
            sequence_str = self.form_vars['sequence_var'].get().strip()
            resources = self.form_vars['resources_var'].get().strip()
            budget_str = self.form_vars['budget_var'].get().strip()
            section_str = self.section_var.get()

            if not all([task, action, duration_str, sequence_str]):
                messagebox.showerror("Error", "Please fill in Task, Action Needed, Duration, and Sequence fields.")
                return

            duration = int(duration_str)
            sequence = int(sequence_str)
            budget = float(budget_str) if budget_str else 0.0

            section = ActivitySection.PRE_KICKOFF if section_str == ActivitySection.PRE_KICKOFF.value else ActivitySection.POST_KICKOFF

            # Create activity
            activity = Activity(
                task=task,
                action_needed=action,
                duration=duration,
                precursor=precursor,
                sequence=sequence,
                resources=resources,
                budget=budget,
                section=section
            )

            self.activities_data.append(activity)

            # Add to treeview
            self.activities_tree.insert("", tk.END, values=(
                section.value, task, action, duration, precursor, sequence, resources, budget
            ))

            # Update activities list count
            self.update_activities_count()

            # Clear form
            for var in self.form_vars.values():
                var.set("")

            messagebox.showinfo("Success", "Activity added successfully!")

        except ValueError as e:
            messagebox.showerror("Error", "Please enter valid numbers for Duration, Sequence, and Budget.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def setup_context_menu(self):
        """Setup right-click context menu for activities tree"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Delete Activity", command=self.delete_activity_with_confirmation)
        
        # Bind right-click event
        self.activities_tree.bind("<Button-3>", self.show_context_menu)  # Right-click
        
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        # Select the item under cursor
        item = self.activities_tree.identify_row(event.y)
        if item:
            self.activities_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def delete_activity_with_confirmation(self):
        """Delete the selected activity with confirmation dialog"""
        selected_item = self.activities_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select an activity to delete.")
            return

        # Get activity details for confirmation
        item_values = self.activities_tree.item(selected_item[0])['values']
        activity_name = item_values[1] if len(item_values) > 1 else "this activity"
        
        # Show confirmation dialog
        confirm = messagebox.askyesno(
            "Confirm Deletion", 
            f"Are you sure you want to delete the activity:\n\n'{activity_name}'?\n\nThis action cannot be undone.",
            icon='warning'
        )
        
        if not confirm:
            return

        # Get the index of the selected item
        index = self.activities_tree.index(selected_item[0])

        # Remove from data and treeview
        if 0 <= index < len(self.activities_data):
            del self.activities_data[index]
            self.activities_tree.delete(selected_item[0])
            self.update_activities_count()
            messagebox.showinfo("Success", "Activity deleted successfully!")

    def clear_all_with_confirmation(self):
        """Clear all activities and reset the form with confirmation"""
        if len(self.activities_data) == 0:
            messagebox.showinfo("Information", "No data to clear.")
            return
            
        confirm = messagebox.askyesno(
            "Confirm Clear All", 
            f"Are you sure you want to clear all project data?\n\n"
            f"This will remove:\n"
            f"• Project title\n"
            f"• All {len(self.activities_data)} activities\n"
            f"• Form data\n\n"
            f"This action cannot be undone.",
            icon='warning'
        )
        
        if confirm:
            self.activities_data.clear()
            self.activities_tree.delete(*self.activities_tree.get_children())
            self.project_title_var.set("")
            for var in self.form_vars.values():
                var.set("")
            self.update_activities_count()
            messagebox.showinfo("Success", "All data cleared!")

    def generate_preview_text(self, project):
        """Generate preview text for both display and file output"""
        schedules = ScheduleCalculator.calculate_schedules(project)
        max_activities = ScheduleCalculator.get_max_duration_activities(project)

        preview_text = f"Project Schedule Summary\n"
        preview_text += f"=" * 50 + "\n\n"
        preview_text += f"Project Title: {project.title}\n"
        preview_text += f"Calendar Format: {project.calendar_format.value}\n"
        preview_text += f"Total Activities: {len(project.activities)}\n"
        
        # Budget summary
        total_budget = sum(activity.budget for activity in project.activities)
        preview_text += f"Total Budget: ${total_budget:,.2f}\n\n"

        for section in [ActivitySection.PRE_KICKOFF, ActivitySection.POST_KICKOFF]:
            activities = project.get_activities_by_section(section)
            if not activities:
                continue
                
            preview_text += f"{section.value} ({len(activities)} activities):\n"
            preview_text += f"-" * 40 + "\n"
            activities.sort(key=lambda x: x.sequence)

            current_sequence = None
            for activity in activities:
                if current_sequence != activity.sequence:
                    current_sequence = activity.sequence
                    raw_schedule = schedules.get(activity.sequence, 0)
                    adjusted_schedule = ScheduleCalculator.apply_calendar_format(raw_schedule, project.calendar_format)
                    preview_text += f"\nSequence {activity.sequence} (Schedule: {adjusted_schedule} days):\n"

                is_max = any(activity is max_act for max_act, _ in max_activities)
                max_indicator = " [MAX DURATION]" if is_max else ""
                
                preview_text += f"  • {activity.task}\n"
                preview_text += f"    Duration: {activity.duration} days{max_indicator}\n"
                preview_text += f"    Action: {activity.action_needed}\n"
                preview_text += f"    Resources: {activity.resources}\n"
                preview_text += f"    Budget: ${activity.budget:,.2f}\n"
                if activity.precursor:
                    preview_text += f"    Precursor: {activity.precursor}\n"
                preview_text += "\n"

        return preview_text

    def preview_schedule(self):
        """Show a preview of the schedule calculations"""
        if not self.validate_project():
            return

        project = self.create_project()
        preview_text = self.generate_preview_text(project)

        # Show preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Schedule Preview")
        preview_window.geometry("600x400")

        text_widget = tk.Text(preview_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", preview_text)
        text_widget.config(state=tk.DISABLED)

        scrollbar = ttk.Scrollbar(preview_window, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)

    def generate_excel(self):
        """Generate and save the Excel file"""
        if not self.validate_project():
            return

        try:
            # Create default filename based on project title
            project_title = self.project_title_var.get().strip()
            safe_title = "".join(c for c in project_title if c.isalnum() or c in (' ', '-', '_')).strip()
            default_filename = f"{safe_title}_schedule.xlsx" if safe_title else "project_schedule.xlsx"

            # Get output file path with improved dialog
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Project Schedule As...",
                initialfile=default_filename
            )

            if not file_path:
                return  # User cancelled

            # Show progress message
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Generating Excel...")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # Center the progress window
            progress_window.geometry("+%d+%d" % (
                self.root.winfo_rootx() + 50,
                self.root.winfo_rooty() + 50
            ))
            
            ttk.Label(progress_window, text="Generating Excel file...", font=('Arial', 12)).pack(expand=True)
            progress_window.update()

                         # Generate Excel file
            project = self.create_project()
            generator = ExcelGenerator(project)
            generator.generate(file_path)
            
            # Generate accompanying .txt file with schedule details
            txt_file_path = file_path.replace('.xlsx', '_details.txt')
            try:
                preview_text = self.generate_preview_text(project)
                with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
                    txt_file.write(preview_text)
            except Exception as txt_error:
                print(f"Warning: Could not create .txt file: {txt_error}")
            
            # Close progress window
            progress_window.destroy()

            # Show success message with option to open file location
            files_created = f"Files created:\n• {file_path}\n• {txt_file_path}"
            result = messagebox.askyesno(
                "Success", 
                f"Project files generated successfully!\n\n"
                f"{files_created}\n\n"
                f"Would you like to open the file location?",
                icon='info'
            )
            
            if result:
                import os
                import subprocess
                import platform
                
                # Open file location in file explorer
                if platform.system() == "Windows":
                    subprocess.run(f'explorer /select,"{file_path}"')
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", "-R", file_path])
                else:  # Linux
                    subprocess.run(["xdg-open", os.path.dirname(file_path)])

        except Exception as e:
            # Close progress window if it exists
            try:
                progress_window.destroy()
            except:
                pass
            messagebox.showerror("Error", f"Failed to generate Excel file:\n{str(e)}")

    def validate_project(self):
        """Validate project data before processing"""
        if not self.project_title_var.get().strip():
            messagebox.showerror("Error", "Please enter a project title.")
            return False

        if not self.activities_data:
            messagebox.showerror("Error", "Please add at least one activity.")
            return False

        return True

    def create_project(self):
        """Create a Project object from the GUI data"""
        # Get calendar format
        calendar_format = CalendarFormat.FIVE_DAY
        for fmt in CalendarFormat:
            if fmt.value == self.calendar_format_var.get():
                calendar_format = fmt
                break

        # Create project
        project = Project(
            title=self.project_title_var.get().strip(),
            calendar_format=calendar_format
        )

        # Add activities
        for activity in self.activities_data:
            project.add_activity(activity)

        return project
    
    def update_activities_count(self):
        """Update the activities list frame title with current count"""
        count = len(self.activities_data)
        self.list_frame.config(text=f"Activities List ({count} activities)")

    def run(self):
        """Start the GUI application"""
        self.root.mainloop()


def main():
    """Main function to run the Project Scheduler application"""
    app = ProjectSchedulerGUI()
    app.run()


if __name__ == "__main__":
    main() 