import pandas as pd
import warnings

warnings.filterwarnings('ignore')

def time_str_to_seconds(time_str):
    if time_str == '0':
        return 0
    hours, minutes = map(int, time_str.split(':'))
    return hours * 3600 + minutes * 60

def seconds_to_hours_rounded(total_seconds):
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    total_minutes = (total_seconds // 60)
    if seconds >= 30:
        total_minutes += 1
    hours = total_minutes // 60
    return hours

def process_attendance(file_path):
    attendance_data = pd.read_excel(file_path, header=None)

    EMP_NAME_X = 3
    EMP_NAME_Y = 4
    PADDING = 6
    DAY_NAME_X = 6
    DAY_NAME_Y = 1
    DATA_X = 7
    DATA_Y = 1
    DEPARTMENT_X = 3
    DEPARTMENT_Y = 10
    PRESENT_DAYS_X = 3
    PRESENT_DAYS_Y = 15
    ABSENT_DAYS_X = 3
    ABSENT_DAYS_Y = 17

    num_days = len(attendance_data.columns) - 1
    num_employees = (len(attendance_data) - 3) // 6

    final_output = pd.DataFrame(columns=[
        'Employee Name', 'Department', 'Present Days', 'Absent Days', 
        'Overtime (hours)', 'Undertime (hours)', 'Adjusted Overtime (hours)'
    ])

    for N in range(1, num_employees + 1):
        total_overtime_seconds = 0
        total_undertime_seconds = 0
        emp_row = EMP_NAME_X + (N - 1) * PADDING
        employee_name = attendance_data.iloc[emp_row, EMP_NAME_Y].strip()
        department = attendance_data.iloc[emp_row, DEPARTMENT_Y].split(":")[1].strip() if pd.notna(attendance_data.iloc[emp_row, DEPARTMENT_Y]) else "Unknown"
        present_days = attendance_data.iloc[emp_row, PRESENT_DAYS_Y].split(":")[1].strip() if pd.notna(attendance_data.iloc[emp_row, PRESENT_DAYS_Y]) else "0"
        absent_days = attendance_data.iloc[emp_row, ABSENT_DAYS_Y].split(":")[1].strip() if pd.notna(attendance_data.iloc[emp_row, ABSENT_DAYS_Y]) else "0"

        for M in range(1, num_days + 1):
            data_row = DATA_X + (N - 1) * PADDING
            data_col = DATA_Y + (M - 1)
            data_string = attendance_data.iloc[data_row, data_col]
            if pd.isna(data_string) or data_string.strip() == '':
                continue

            try:
                in_time_str, out_time_str = data_string.strip().split('\n')[:2]
                in_seconds = time_str_to_seconds(in_time_str)
                out_seconds = time_str_to_seconds(out_time_str)
                work_duration_seconds = out_seconds - in_seconds

                day_name_row = DAY_NAME_X + (N - 1) * PADDING
                day_name_col = DAY_NAME_Y + (M - 1)
                day_name = attendance_data.iloc[day_name_row, day_name_col].strip()

                if day_name == "Su":
                    total_overtime_seconds += work_duration_seconds
                else:
                    if work_duration_seconds > 8 * 3600:
                        total_overtime_seconds += work_duration_seconds - 8 * 3600
                    elif work_duration_seconds < 8 * 3600:
                        total_undertime_seconds += 8 * 3600 - work_duration_seconds

            except Exception as e:
                print(f"Data format error for employee {employee_name} on day {M}: {e}")
                continue

        adjusted_overtime_seconds = total_overtime_seconds - total_undertime_seconds
        overtime_hours = seconds_to_hours_rounded(total_overtime_seconds)
        undertime_hours = seconds_to_hours_rounded(total_undertime_seconds)
        adjusted_overtime_hours = seconds_to_hours_rounded(abs(adjusted_overtime_seconds))
        if adjusted_overtime_seconds < 0:
            adjusted_overtime_hours = -adjusted_overtime_hours
        overtime_str = str(overtime_hours)
        undertime_str = str(undertime_hours)
        adjusted_overtime_str = str(adjusted_overtime_hours)

        final_output = final_output.append({
            'Employee Name': employee_name,
            'Department': department,
            'Present Days': int(float(present_days)),
            'Absent Days': int(float(absent_days)),
            'Overtime (hours)': overtime_str,
            'Undertime (hours)': undertime_str,
            'Adjusted Overtime (hours)': adjusted_overtime_str
        }, ignore_index=True)

    return final_output