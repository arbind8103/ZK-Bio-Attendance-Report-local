import requests
import pandas as pd
from datetime import datetime, timedelta
from config import API_URL, EMP_API_URL, AUTH

def get_month_range():
    now = datetime.now()
    if now.month == 1:
        start = datetime(now.year - 1, 12, 26)
    else:
        start = datetime(now.year, now.month - 1, 26)
    end = now.replace(hour=23, minute=59, second=59, microsecond=0)
    return start.date(), end.date()

def fetch_all_employee_data():
    emp_data = []
    page = 1
    while True:
        try:
            response = requests.get(f"{EMP_API_URL}?page={page}", auth=AUTH, timeout=15)
            response.raise_for_status()
            result = response.json()
            data_page = result.get("data", [])
            if not data_page:
                break
            emp_data.extend(data_page)
            if not result.get("next"):
                break
            page += 1
        except requests.exceptions.RequestException as e:
            print(f"❌ Employee API error on page {page}: {e}")
            break

    emp_df = pd.json_normalize(emp_data)
    if emp_df.empty:
        return pd.DataFrame()

    emp_df = emp_df[['emp_code', 'first_name', 'enroll_sn', 'department.dept_name', 'area']]
    emp_df['area_code'] = emp_df['area'].apply(lambda x: x[0]['area_code'] if x else None)
    emp_df['area_name'] = emp_df['area'].apply(lambda x: x[0]['area_name'] if x else None)
    emp_df.rename(columns={'first_name': 'emp_name', 'department.dept_name': 'dept_name'}, inplace=True)
    emp_df.drop(columns=['area'], inplace=True)
    return emp_df

def calculate_worked_hours(row):
    punch_in = row['punch_in']
    punch_out = row['punch_out']
    total = (punch_out - punch_in).total_seconds() / 3600
    lunch_start = punch_in.replace(hour=13, minute=0)
    lunch_end = lunch_start + timedelta(minutes=45)
    if punch_out > lunch_start and punch_in < lunch_end:
        total -= 0.75
    return round(total, 2)

def fetch_attendance_data():
    start_date, end_date = get_month_range()
    print(f"📥 Fetching data from {start_date} to {end_date}")

    all_logs = []
    page = 1

    while True:
        url = f"{API_URL}?start_time={start_date} 00:00:00&end_time={end_date} 23:59:59&page={page}"
        try:
            response = requests.get(url, auth=AUTH, timeout=15)
            response.raise_for_status()
            data = response.json()
            logs = data.get('data', [])
            if not logs:
                break
            all_logs.extend(logs)
            if not data.get("next"):
                break
            page += 1
        except requests.exceptions.RequestException as e:
            print(f"❌ Failed to fetch data on page {page}: {e}")
            break

    if not all_logs:
        print("No punch data found.")
        return pd.DataFrame()

    df = pd.DataFrame(all_logs)
    df['punch_time'] = pd.to_datetime(df['punch_time'])
    df['date'] = df['punch_time'].dt.date

    first_punch = df.sort_values('punch_time').groupby(['emp_code', 'date']).first().reset_index()
    last_punch = df.sort_values('punch_time').groupby(['emp_code', 'date']).last().reset_index()

    first_punch = first_punch[['emp_code', 'date', 'punch_time']]
    last_punch = last_punch[['emp_code', 'date', 'punch_time']].rename(columns={'punch_time': 'punch_out'})

    final_df = pd.merge(first_punch, last_punch, on=['emp_code', 'date'])
    final_df.rename(columns={'punch_time': 'punch_in'}, inplace=True)
    final_df['worked_hours'] = final_df.apply(calculate_worked_hours, axis=1)

    emp_df = fetch_all_employee_data()
    if emp_df.empty:
        print("❌ Employee data could not be fetched.")
        return pd.DataFrame()

    full_dates = pd.date_range(start=start_date, end=end_date).date
    all_combinations = pd.MultiIndex.from_product([emp_df['emp_code'], full_dates], names=['emp_code', 'date'])
    base_df = pd.DataFrame(index=all_combinations).reset_index()

    merged_df = pd.merge(base_df, final_df, on=['emp_code', 'date'], how='left')
    merged_df = pd.merge(merged_df, emp_df, on='emp_code', how='left')

    return merged_df
