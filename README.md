🕐 Attendance Processor
A powerful Streamlit-based automation tool designed to process employee attendance logs from Excel files. It calculates total hours worked, identifies missing punches, compares attendance against customizable weekly targets, and generates a professionally formatted multi-sheet Excel report.

✨ Features
Automated Hour Calculation: Converts raw "In/Out" punch strings into precise decimal and HH.MM durations.

Intelligent Week Numbering: Uses a custom "Calendar-Style" week logic where Week 1 always starts on the 1st of the month and increments every Monday.

Missing Punch Detection: Automatically identifies days with single punches and provides a UI to "fix" them before generating the report.

Dynamic Targets: Calculates fair weekly targets based on the number of non-Sunday working days in that specific relative week.

Professional Reporting: Generates a single .xlsx file containing:

Consolidated Report: Monthly totals for all employees.

Weekly Summary: Performance vs. Target for every week.

Individual Sheets: A dedicated tab for every employee with a daily breakdown.

🛠️ Installation
Clone the repository:

Bash
git clone https://github.com/yourusername/attendance-processor.git
cd attendance-processor
Install dependencies:
Make sure you have Python 3.8+ installed.

Bash
pip install streamlit openpyxl
Run the application:

Bash
streamlit run app.py
📋 How to Use
Launch the App: Run the streamlit command above to open the interface in your browser.

Upload Data: Upload an .xlsx file containing a sheet named Logs (formatted with employee IDs and punch strings).

Set Targets: Use the sidebar to adjust the Weekly Target Hours (default is 51.0).

Fix Punches: If the app detects missing clock-ins or clock-outs, enter the estimated times in the provided text boxes.

Generate & Download: Click "Generate Excel Report" to process the data and download your formatted summary.

🧮 Logic Breakdown
Custom Week Calculation
Unlike the standard ISO week logic which often starts Week 1 in the previous month, this tool uses a "Month-Centric" approach:

Week 1 always begins on the 1st day of the month.

The week number increments only when a Monday is encountered.

This ensures that payroll and attendance summaries align perfectly with the calendar month being viewed.

Target Calculations
The tool calculates a daily_target based on your weekly input.

Example: If the Weekly Target is 51.0 hours for a 6-day week, the daily target is 8.5 hours. If Week 1 only has 3 working days (e.g., Thu, Fri, Sat), the target for that specific week is automatically set to 25.5 hours.

📦 Project Structure
app.py: The main Streamlit application code.

requirements.txt: List of Python dependencies.

README.md: This documentation.

🛠️ Technologies Used
Python

Streamlit (Web Interface)

Openpyxl (Excel Engine)
