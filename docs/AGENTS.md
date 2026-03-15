oSH_THEME="robbyrussell"
nstructions

## Goal
Implement a Python attendance-report script for this project.

## Files
- `a.xls`: raw attendance data
- `demo1.xlsx`: target output format reference
- `task.md`: full business rules

## Requirements
- First inspect `a.xls` and `demo1.xlsx`
- Then create `attendance_report.py`
- Output file should be `result.xlsx`
- Keep all adjustable column mappings in a config section
- Prefer pandas + openpyxl + xlrd
- Code must be runnable, not pseudocode
- After coding, run the script and fix obvious errors if possible

## Notes
- Treat employee ID as primary key
- Only count workdays
- Follow `task.md` exactly for attendance status mapping
