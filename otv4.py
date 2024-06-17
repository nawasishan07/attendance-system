import pandas as pd
import warnings

warnings.filterwarnings('ignore')

def time_str_to_seconds(time_str):
    if time_str == '0':
        return 0
    hours, minutes = map(int, time_str.split(':'))
    return hours * 3600 + minutes * 60

# def seconds_to_HHMM(total_seconds):
#     # Function to convert seconds to HH:MM format
#     hours = total_seconds // 3600
#     minutes = (total_seconds % 3600) // 60
#     return f"{hours:02d}:{minutes:02d}"

def seconds_to_hours_rounded(total_seconds):
    # Convert total seconds to total hours, rounding up or down based on seconds
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    total_minutes = (total_seconds // 60)

    # Round up if seconds are 30 or more
    if seconds >= 30:
        total_minutes += 1

    # Convert minutes to hours
    hours = total_minutes // 60
    return hours

file_path = "SEPT U I ATTENDANCE SUMMARY.xls"
attendance_data = pd.read_excel(file_path, header=None)

# Initializing the constants
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

# Placeholder for the final output DataFrame
final_output = pd.DataFrame(columns=[
    'Employee Name', 'Department', 'Present Days', 'Absent Days', 
    'Overtime (hours)', 'Undertime (hours)', 'Adjusted Overtime (hours)'
])

# Loop through each employee
for N in range(1, num_employees + 1):
    # Initialize total overtime and undertime for the employee in seconds
    total_overtime_seconds = 0
    total_undertime_seconds = 0
    
    # Calculate the row for the employee's data and retrieve it
    emp_row = EMP_NAME_X + (N - 1) * PADDING
    employee_name = attendance_data.iloc[emp_row, EMP_NAME_Y].strip()
    department = attendance_data.iloc[emp_row, DEPARTMENT_Y].split(":")[1].strip() if pd.notna(attendance_data.iloc[emp_row, DEPARTMENT_Y]) else "Unknown"
    present_days = attendance_data.iloc[emp_row, PRESENT_DAYS_Y].split(":")[1].strip() if pd.notna(attendance_data.iloc[emp_row, PRESENT_DAYS_Y]) else "0"
    absent_days = attendance_data.iloc[emp_row, ABSENT_DAYS_Y].split(":")[1].strip() if pd.notna(attendance_data.iloc[emp_row, ABSENT_DAYS_Y]) else "0"

    # Loop through each day for the current employee
    for M in range(1, num_days + 1):
        # Calculate the row and column for the day's data
        data_row = DATA_X + (N - 1) * PADDING
        data_col = DATA_Y + (M - 1)

        # Fetch the in and out times
        data_string = attendance_data.iloc[data_row, data_col]
        if pd.isna(data_string) or data_string.strip() == '':
            # If data is missing, we skip this day
            continue

        try:
            in_time_str, out_time_str = data_string.strip().split('\n')[:2]
            in_seconds = time_str_to_seconds(in_time_str)
            out_seconds = time_str_to_seconds(out_time_str)

            # Calculate work duration in seconds
            work_duration_seconds = out_seconds - in_seconds

            # Check if the current day is a Sunday
            day_name_row = DAY_NAME_X + (N - 1) * PADDING
            day_name_col = DAY_NAME_Y + (M - 1)
            day_name = attendance_data.iloc[day_name_row, day_name_col].strip()

            if day_name == "Su":
                total_overtime_seconds += work_duration_seconds
            else:
                # Calculate overtime and undertime in seconds
                if work_duration_seconds > 8 * 3600:  # Greater than 8 hours in seconds
                    total_overtime_seconds += work_duration_seconds - 8 * 3600
                elif work_duration_seconds < 8 * 3600:  # Less than 8 hours in seconds
                    total_undertime_seconds += 8 * 3600 - work_duration_seconds

        except Exception as e:
            print(f"Data format error for employee {employee_name} on day {M}: {e}")
            continue

    # Calculate Adjusted Overtime in seconds
    adjusted_overtime_seconds = total_overtime_seconds - total_undertime_seconds

    # Format total overtime, undertime, and adjusted overtime into "HH:MM"
    # overtime_str = seconds_to_HHMM(total_overtime_seconds)
    # undertime_str = seconds_to_HHMM(total_undertime_seconds)
    # adjusted_overtime_str = seconds_to_HHMM(abs(adjusted_overtime_seconds))
    # if adjusted_overtime_seconds < 0:
    #     adjusted_overtime_str = "-" + adjusted_overtime_str

    # Format total overtime, undertime, and adjusted overtime into "HH:MM" with rounding
    overtime_hours = seconds_to_hours_rounded(total_overtime_seconds)
    undertime_hours = seconds_to_hours_rounded(total_undertime_seconds)
    adjusted_overtime_hours = seconds_to_hours_rounded(abs(adjusted_overtime_seconds))
    if adjusted_overtime_seconds < 0:
        adjusted_overtime_hours = -adjusted_overtime_hours
    overtime_str = str(overtime_hours)
    undertime_str = str(undertime_hours)
    adjusted_overtime_str = str(adjusted_overtime_hours)

    # Add the data to the final DataFrame
    final_output = final_output.append({
        'Employee Name': employee_name,
        'Department': department,
        'Present Days': int(float(present_days)),
        'Absent Days': int(float(absent_days)),
        'Overtime (hours)': overtime_str,
        'Undertime (hours)': undertime_str,
        'Adjusted Overtime (hours)': adjusted_overtime_str
    }, ignore_index=True)

# Define the output file path
output_file_path = 'final_overtime_report.xlsx'

# Export the DataFrame to an Excel file
final_output.to_excel(output_file_path, index=False)

# This would be the output statement if the code was executed
print(f"Overtime report generated: {output_file_path}")