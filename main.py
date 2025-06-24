import pandas as pd
from generate_reports import generate_attendance_reports
from utils import fetch_attendance_data

if __name__ == "__main__":
    # Fetch attendance data merged with employee info
    emp_df = fetch_attendance_data()

    # Ensure correct data types
    emp_df['punch_in'] = emp_df['punch_in'].astype(str).replace("nan", None)
    emp_df['punch_out'] = emp_df['punch_out'].astype(str).replace("nan", None)
    emp_df['worked_hours'] = emp_df['worked_hours'].fillna(0)

    # Generate master Excel and PDF reports and send to HR
    generate_attendance_reports(emp_df)
