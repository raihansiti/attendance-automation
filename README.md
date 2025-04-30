# Attendance Automation with Google Apps Script
This project automates the matching of worker attendance records with corresponding agents using Google Sheets and Google Apps Script. It replaces manual processes with accurate data linking, visual highlights, and a dynamic summary legend for quick reference.

## Features
- Relational Matching: Matches Worker ID from AttendanceList with the agent in AgentList using unique identifiers.
- Color-coded Highlights: Each agent is assigned a unique soft color to visually differentiate matched attendance rows.
- Auto-Fill Agent Name: Automatically inserts the matched Agent Name into the attendance sheet for traceability.
- Custom UI Integration: Adds a menu item (“Attendance Tools”) for easy execution.
- Legend Sheet: Generates a clean and user-friendly legend sheet to show which color belongs to which agent.
- Error Handling: Alerts the user when required sheets or data are missing.
- Row Control: Only applies changes within specified rows to avoid unwanted data overwrite.

## Technologies Used
- Google Apps Script (JavaScript-based)
- Google Sheets

## How It Works
1. User clicks Attendance Tools > Match and Highlight Attendance.
2. The script:
   - Reads Agent IDs and Names from AgentList.
   - Reads multiple columns of Worker IDs from AttendanceList.
   - Matches entries and applies color coding for visual grouping.
   - Fills agent name into Column AP (column 42).
   - Writes a color legend into a new or existing Results sheet.

## Why This Project?
This was built to solve a real-world problem where worker attendance was tracked manually. By applying database design principles (like entity relationships and unique IDs) and using automation, this project saves time, reduces errors, and improves data clarity.

## Sample Screenshot (optional)
*You can include a screenshot of your Google Sheet before/after running the script here.*

## Author
[LinkedIn](https://linkedin.com/in/snraihan96)
