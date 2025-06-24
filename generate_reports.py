import os
import sys
import pandas as pd
from datetime import datetime
from fpdf import FPDF

MASTER_PDF_PATH = "Master_Attendance_Report.pdf"
MASTER_EXCEL_PATH = "Master_Attendance_Report.xlsx"

# ✅ Set LOGO_PATH dynamically (works in PyInstaller build)
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOGO_PATH = os.path.join(BASE_DIR, "logo.png")

def determine_day_status(row):
    weekday = pd.to_datetime(row['date']).weekday()
    if weekday == 6:  # Sunday
        return 'OFF'

    if pd.isna(row['punch_in']) and pd.isna(row['punch_out']):
        return 'A'

    if pd.notna(row['worked_hours']) and row['worked_hours'] >= 8:
        return 'P'

    if pd.notna(row['punch_in']):
        try:
            punch_in_time = pd.to_datetime(row['punch_in'])
            if pd.isna(punch_in_time):
                return 'A'
            punch_in_time = punch_in_time.time()
            cutoff_time = datetime.strptime('09:15', '%H:%M').time()
            if punch_in_time > cutoff_time:
                return 'HD' if row['worked_hours'] < 8 else 'P'
            else:
                return 'P'
        except Exception as e:
            print(f"⚠️ Error parsing punch_in for {row['emp_code']} on {row['date']}: {e}")
            return 'A'
    return 'A'

def generate_attendance_reports(emp_df):
    emp_df['day_status'] = emp_df.apply(determine_day_status, axis=1)
    emp_df['date'] = pd.to_datetime(emp_df['date'])

    summary_data = []
    excel_writer = pd.ExcelWriter(MASTER_EXCEL_PATH, engine='xlsxwriter')
    workbook = excel_writer.book

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)

    grouped = emp_df.groupby('emp_code')
    for emp_code, data in grouped:
        emp_name = data['emp_name'].iloc[0]
        dept_name = data['dept_name'].iloc[0] if 'dept_name' in data else 'N/A'

        df = data[['date', 'punch_in', 'punch_out', 'worked_hours', 'day_status']].copy()
        df.columns = ['Date', 'Punch In', 'Punch Out', 'Worked Hours', 'Status']
        df.sort_values('Date', inplace=True)

        # Excel Sheet
        sheet_name = emp_name[:31]
        df.to_excel(excel_writer, sheet_name=sheet_name, index=False, startrow=8)
        worksheet = excel_writer.sheets[sheet_name]
        worksheet.set_column("A:E", 20)

        title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        subtitle = workbook.add_format({'italic': True, 'align': 'center'})
        label = workbook.add_format({'bold': True, 'font_color': 'navy'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2'})

        worksheet.merge_range('A1:E1', 'Supreme Automobile SARL', title)
        worksheet.merge_range('A2:E2', 'No. 5406, 12ème Rue Industriel C/Limete, Kinshasa, DRC', subtitle)
        worksheet.write("A4", "Employee ID", label)
        worksheet.write("B4", emp_code)
        worksheet.write("A5", "Employee Name", label)
        worksheet.write("B5", emp_name)
        worksheet.write("A6", "Department", label)
        worksheet.write("B6", dept_name)

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(8, col_num, value, header_format)

        format_P = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        format_HD = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
        format_A = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        status_col = 'E'
        last_row = len(df) + 8
        worksheet.conditional_format(f'{status_col}9:{status_col}{last_row}', {'type': 'cell', 'criteria': '==', 'value': '"P"', 'format': format_P})
        worksheet.conditional_format(f'{status_col}9:{status_col}{last_row}', {'type': 'cell', 'criteria': '==', 'value': '"HD"', 'format': format_HD})
        worksheet.conditional_format(f'{status_col}9:{status_col}{last_row}', {'type': 'cell', 'criteria': '==', 'value': '"A"', 'format': format_A})

        # PDF section
        pdf.add_page()
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=10, y=8, w=30)
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, 'Supreme Automobile SARL', ln=True, align='C')
        pdf.set_font("Arial", '', 10)
        pdf.cell(0, 8, 'No. 5406, 12ème Rue Industriel C/Limete, Kinshasa, DRC', ln=True, align='C')
        pdf.ln(8)
        pdf.set_font("Arial", '', 11)
        pdf.cell(0, 6, f"Employee ID: {emp_code}", ln=True)
        pdf.cell(0, 6, f"Employee Name: {emp_name}", ln=True)
        pdf.cell(0, 6, f"Department: {dept_name}", ln=True)
        pdf.ln(6)

        col_widths = [32, 40, 40, 40, 25]
        headers = ['Date', 'Punch In', 'Punch Out', 'Worked Hours', 'Status']
        pdf.set_font("Arial", 'B', 10)
        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 8, h, border=1, align='C')
        pdf.ln()

        pdf.set_font("Arial", '', 10)
        for _, row in df.iterrows():
            pdf.cell(col_widths[0], 7, str(row['Date'].date()), border=1)
            pdf.cell(col_widths[1], 7, str(row['Punch In']), border=1)
            pdf.cell(col_widths[2], 7, str(row['Punch Out']), border=1)
            pdf.cell(col_widths[3], 7, str(row['Worked Hours']), border=1)
            pdf.cell(col_widths[4], 7, str(row['Status']), border=1, align='C')
            pdf.ln()

        present = (df['Status'] == 'P').sum()
        half_day = (df['Status'] == 'HD').sum()
        absent = (df['Status'] == 'A').sum()
        total_hours = df['Worked Hours'].sum()

        summary_data.append([emp_code, emp_name, dept_name, present, half_day, absent, round(total_hours, 2)])

    # Summary sheet
    summary_df = pd.DataFrame(summary_data, columns=["Emp Code", "Emp Name", "Department", "Present", "Half Day", "Absent", "Total Hours"])
    summary_df.to_excel(excel_writer, sheet_name="Summary", index=False)
    excel_writer.close()

    pdf.output(MASTER_PDF_PATH)
    print(f"✅ PDF saved: {MASTER_PDF_PATH}")
    print(f"✅ Excel saved: {MASTER_EXCEL_PATH}")
