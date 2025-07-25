# Project Scheduler

A Python application that creates Excel project schedules from user inputs through a form-like GUI interface. The application implements advanced scheduling logic with automatic calculations, cell merging, and conditional formatting.

## Features

- **User-friendly GUI**: Form-based interface with enhanced layout and right-click context menus
- **Advanced Excel Generation**: Professional formatting with merged cells, conditional highlighting, and budget totals
- **Smart Scheduling**: Automatic calculation of project timelines with sequence-based dependencies
- **Calendar Format Options**: Support for 5-day, 6-day, and 7-day work weeks
- **Two Project Phases**: Pre-Kickoff and Post Kick-off Activities sections
- **Data Validation**: Input validation and error handling with confirmation dialogs
- **Preview Function**: Preview schedules before generating Excel files
- **Enhanced File Management**: Progress indicators, default filenames, and file location opening

## System Rules

The application follows these key business rules:

1. **Cell Merging**: Cells in the "Schedule (in days)" column are merged for rows sharing the same sequence value
2. **Schedule Calculation**: 
   - Pre-Kickoff Activities always show 0 in the Schedule column (scheduling starts post-kickoff)
   - Post Kick-off Activities calculate incrementally based on maximum duration per sequence
3. **Duration Highlighting**: Maximum duration values within each sequence group are highlighted with red text
4. **Calendar Adjustments**: Automatic weekend/non-working day calculations based on selected format
5. **Budget Total**: Automatic calculation and display of total project budget with highlighting
6. **Dual File Output**: Generates both Excel (.xlsx) and detailed text (.txt) files simultaneously

## Installation

1. **Clone or download** this repository to your local machine

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python project_scheduler.py
   ```

## Usage

### Starting the Application

Run the main script:
```bash
python project_scheduler.py
```

### Using the GUI

1. **Project Setup**:
   - Enter your project title in the "Project Title" field
   - Select your preferred calendar format (5-day, 6-day, or 7-day week)

2. **Adding Activities**:
   - Fill in the activity form with the following fields:
     - **Task**: Description of the activity/task
     - **Action Needed**: Required action for this task
     - **Duration (days)**: Number of working days needed
     - **Precursor**: Dependencies or previous tasks
     - **Sequence**: Sequence number for grouping (activities with same sequence run in parallel)
     - **Resources**: Required resources
     - **Budget**: Associated budget (optional)
     - **Section**: Choose between "Pre-Kickoff Activities" or "Post Kick-off Activities"
   - Click "Add Activity" to add the activity to your project

3. **Managing Activities**:
   - View all added activities in the Activities List table (shows activity count)
   - **Delete activities**: Use "Delete Selected" button OR right-click for context menu
   - **Confirmation dialogs**: Get detailed confirmation before deleting or clearing
   - Use "Clear All" to reset the entire project with confirmation

4. **Generating Output**:
   - **Preview Schedule**: See a text preview of your schedule calculations
   - **Generate Excel**: Create and save both Excel (.xlsx) and text (.txt) files with:
     - Progress indicator during generation
     - Default filename based on project title
     - Detailed text file with complete project breakdown
     - Option to open file location after creation

### Column Headers in Excel Output

The generated Excel file includes these columns:
- **Activities/Tasks**: Task descriptions
- **Action Needed**: Required actions
- **Duration**: Duration in days
- **Precursor**: Dependencies
- **Sequence**: Sequence grouping numbers
- **Schedule (in days)**: Calculated schedule dates (merged for same sequences, 0 for Pre-Kickoff)
- **Resources**: Required resources
- **Budget**: Associated costs (includes highlighted total row)

## Example Usage

### Sample Project Data

Here's an example of how to structure your project:

**Project Title**: "Website Development Project"

**Pre-Kickoff Activities** (Sequence 1):
- Task: "Requirements Gathering", Duration: 3 days
- Task: "Budget Approval", Duration: 2 days
- Task: "Team Assembly", Duration: 1 day

**Post Kick-off Activities**:
- **Sequence 1** (Development Phase):
  - Task: "Frontend Development", Duration: 10 days
  - Task: "Backend Development", Duration: 12 days
  - Task: "Database Setup", Duration: 8 days

- **Sequence 2** (Testing Phase):
  - Task: "Unit Testing", Duration: 5 days
  - Task: "Integration Testing", Duration: 7 days

**Schedule Calculation Result**:
- Pre-Kickoff: All activities show 0 days (scheduling starts post-kickoff)
- Sequence 1: Shows 12 days (maximum of 10, 12, 8)
- Sequence 2: Shows 19 days (12 + 7, where 7 is max of 5, 7)

### Calendar Format Examples

- **5-day week**: 12 working days → 16 calendar days (includes weekends)
- **6-day week**: 12 working days → 14 calendar days (includes Sundays)
- **7-day week**: 12 working days → 12 calendar days (no adjustments)

## File Structure

```
PSP/
├── project_scheduler.py     # Main application file
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Requirements

- Python 3.7+
- openpyxl 3.1.2+
- tkinter (usually included with Python)

## Troubleshooting

### Common Issues

1. **"Module not found" errors**: Ensure you've installed all dependencies with `pip install -r requirements.txt`

2. **Excel file won't open**: Make sure you have Excel or a compatible spreadsheet application installed

3. **GUI doesn't appear**: Ensure tkinter is properly installed with your Python distribution

4. **Permission errors when saving**: Choose a location where you have write permissions

### Input Validation

The application validates:
- Required fields (Task, Action Needed, Duration, Sequence)
- Numeric values for Duration, Sequence, and Budget
- Positive duration values

## Advanced Features

### Cell Merging Logic
The application automatically merges cells in the "Schedule (in days)" column for activities that share the same sequence number, creating a clean, professional appearance.

### Duration Highlighting
Activities with the maximum duration within each sequence group are highlighted with bold red text for easy identification of critical path items.

### Dual File Generation
The application automatically creates two files when generating output:
- **Excel file (.xlsx)**: Professional formatted spreadsheet with merged cells, calculations, and formatting
- **Text file (.txt)**: Detailed project breakdown with schedule summary, activity details, and budget information

### Multiple Subsections
You can create multiple subsections within the "Post Kick-off Activities" by using different sequence numbers, all maintained within the same worksheet.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify all requirements are met
3. Ensure input data follows the expected format

## License

This project is provided as-is for educational and practical use. 