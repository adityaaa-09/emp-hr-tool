import pandas as pd
import json
from datetime import datetime, timedelta
import os
import re
import datetime

## Sub Function (NESTED)
def update_days_from_filename(employee_dict, file_name):
    """
    Updates the 'Days' list in the employee dictionary by adding proper date information
    based on the file name (which contains month and year).

    Args:
        employee_dict (dict): Dictionary containing employee attendance data
        file_name (str): Name of the CSV file (e.g. "aug_2024_biometric.csv")

    Returns:
        dict: Updated employee dictionary with complete date information
    """
    # Extract month and year from filename
    match = re.match(r'([a-zA-Z]+)_(\d{4})_biometric\.csv', file_name)
    if not match:
        print(f"Warning: Could not extract month/year from filename: {file_name}")
        return employee_dict

    month_abbr = match.group(1).lower()
    year = match.group(2)

    # Map month abbreviations to month numbers
    month_map = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
    }

    month_num = month_map.get(month_abbr, None)
    if month_num is None:
        print(f"Warning: Unknown month abbreviation: {month_abbr}")
        return employee_dict

    # Update employee data in-place
    for employee, data in employee_dict.items():
        days_list = data['Days']
        updated_days = []

        for day_str in days_list:
            if isinstance(day_str, str) and len(day_str) > 0:
                # Extract day number and weekday abbreviation
                parts = day_str.split(maxsplit=1)
                if len(parts) == 2:
                    try:
                        day_num = int(parts[0])
                        day_abbr = parts[1]

                        # Create full date string with day of week
                        date_obj = datetime.date(int(year), month_num, day_num)
                        weekday = date_obj.strftime("%A")  # Full weekday name

                        # Keep the original format but add full date information
                        updated_days.append(f"{day_num} {day_abbr} ({date_obj.strftime('%d %B %Y')}, {weekday})")
                    except (ValueError, IndexError):
                        # If parsing fails, keep the original
                        updated_days.append(day_str)
                else:
                    updated_days.append(day_str)
            else:
                updated_days.append(day_str)

        # Update the 'Days' list in-place
        data['Days'] = updated_days

    return employee_dict


def process_attendance_data(input_csv):
    """
    Extracts 'Status', 'InTime', and 'OutTime' for each employee,
    ensuring dates are correctly aligned.
    """
    # Read CSV without headers to analyze structure
    df = pd.read_csv(input_csv, header=None, encoding="utf-8")

    # Locate the row containing actual dates (assumed to be right above "Employee:")
    date_row_index = df[df.iloc[:, 0] == "Days"].index[0]  # Locate where dates are stored
    dates = df.iloc[date_row_index, 2:].tolist()  # Extract dates from columns (ignoring first two columns)

    # Locate where employee data starts
    emp_rows = df[df.iloc[:, 0] == "Employee:"].index.tolist()

    cleaned_data = [["Employee ID", "Employee Name", "Type"] + dates]  # Update header

    for i in range(len(emp_rows)):
        start_idx = emp_rows[i]
        end_idx = emp_rows[i + 1] if i + 1 < len(emp_rows) else len(df)

        # Extract Employee Name and ID
        emp_info = str(df.iloc[start_idx, 3])
        if ":" in emp_info:
            parts = emp_info.split(":")
            employee_id = parts[0].strip()
            employee_name = parts[-1].strip()
        else:
            employee_id = ""
            employee_name = emp_info.strip()

        # Extract the subset of data belonging to this employee
        emp_data = df.iloc[start_idx:end_idx]

        # Locate 'Status', 'InTime', and 'OutTime' rows
        status_row = emp_data[emp_data.iloc[:, 0] == "Status"]
        intime_row = emp_data[emp_data.iloc[:, 0] == "InTime"]
        outtime_row = emp_data[emp_data.iloc[:, 0] == "OutTime"]

        if not status_row.empty:
            status_values = status_row.iloc[:, 2:].values.flatten()
            cleaned_data.append([employee_id, employee_name, "Status"] + list(status_values))

        if not intime_row.empty:
            intime_values = intime_row.iloc[:, 2:].values.flatten()
            cleaned_data.append([employee_id, employee_name, "InTime"] + list(intime_values))

        if not outtime_row.empty:
            outtime_values = outtime_row.iloc[:, 2:].values.flatten()
            cleaned_data.append([employee_id, employee_name, "OutTime"] + list(outtime_values))

    # Convert cleaned data into DataFrame
    new_df = pd.DataFrame(cleaned_data)
    new_df = new_df.dropna(axis=1, how='all')

    return new_df


def extract_month_year_from_filename(filename):
    """
    Extracts month and year from filenames like 'jan_2024_biometric.csv'
    Returns a tuple of (month_name, year, month_year_key)
    """
    pattern = r'([a-zA-Z]+)_(\d{4})_biometric\.csv'
    match = re.match(pattern, filename)

    if match:
        month_name = match.group(1).lower()
        year = match.group(2)
        month_year_key = f"{month_name}_{year}"
        return (month_name, year, month_year_key)

    return None


def create_employee_dict(df):
    # Extract header row (dates start from column index 3 onward)
    dates = df.iloc[0, 3:].tolist()

    # Initialize the dictionary
    employee_data = {}

    # Start processing from row 1
    i = 1
    while i < len(df):
        employee_id = str(df.iloc[i, 0])
        employee_name = df.iloc[i, 1]
        row_type = df.iloc[i, 2]

        if pd.notna(employee_name) and row_type == "Status":
            # Status row
            status = [str(x) if pd.notna(x) else "NaT" for x in df.iloc[i, 3:].tolist()]

            # InTime and OutTime (assume they exist next)
            in_time = [str(x) if pd.notna(x) else "NaT" for x in df.iloc[i + 1, 3:].tolist()] if i + 1 < len(df) else []
            out_time = [str(x) if pd.notna(x) else "NaT" for x in df.iloc[i + 2, 3:].tolist()] if i + 2 < len(df) else []

            # Construct nested dictionary
            employee_data[employee_name] = {
                "employee_id": employee_id,
                "Days": dates,
                "Status": status,
                "InTime": in_time,
                "OutTime": out_time
            }

            i += 3  # Move to next employee's data
        else:
            i += 1  # Skip irrelevant rows

    return employee_data


def process_attendance_file(csv_file_path):
    """
    Processes a single CSV file and returns the employee attendance dictionary.

    Args:
        csv_file_path (str): Path to the CSV file

    Returns:
        dict: Dictionary containing employee attendance data with updated days
    """
    # Extract the filename from the path
    csv_file = os.path.basename(csv_file_path)

    # Extract month and year information from filename
    file_info = extract_month_year_from_filename(csv_file)

    if not file_info:
        print(f"Warning: Could not extract month/year from filename: {csv_file}")
        return {}

    month_name, year, month_year_key = file_info

    # Process the file and create employee dictionary
    processed_df = process_attendance_data(csv_file_path)
    employee_data = create_employee_dict(processed_df)

    # Update days based on filename
    employee_data = update_days_from_filename(employee_data, csv_file)


    return employee_data


def date_cleaning(employee_dictionary):
    """
    Cleans the date format in employee dictionary by extracting the proper date format
    from strings like '23 F (23 August 2024, Friday)' to '23 August 2024, Friday'

    Args:
        employee_dictionary (dict): Dictionary containing employee attendance data

    Returns:
        dict: Updated dictionary with cleaned date format
    """
    # Create a copy to avoid modifying the original dictionary during iteration
    cleaned_dictionary = {}

    # Process each employee in the dictionary
    for employee_name, employee_data in employee_dictionary.items():
        # Create a copy of the employee data
        cleaned_employee_data = {
            'Status': employee_data['Status'].copy(),
            'InTime': employee_data['InTime'].copy(),
            'OutTime': employee_data['OutTime'].copy(),
            'EmployeeID': employee_data['employee_id']
        }

        # Clean the date format in the Days list
        days_list = employee_data['Days']
        cleaned_days = []

        for day_str in days_list:
            if isinstance(day_str, str) and '(' in day_str and ')' in day_str:
                # Extract the content between parentheses
                start_idx = day_str.find('(') + 1
                end_idx = day_str.find(')')
                if start_idx < end_idx:
                    # Get the clean date format
                    clean_date = day_str[start_idx:end_idx]
                    cleaned_days.append(clean_date)
                else:
                    # If the format is unexpected, keep the original
                    cleaned_days.append(day_str)
            else:
                # If there are no parentheses, keep the original
                cleaned_days.append(day_str)

        # Update the Days list with cleaned dates
        cleaned_employee_data['Days'] = cleaned_days

        # Add the updated employee data to the cleaned dictionary
        cleaned_dictionary[employee_name] = cleaned_employee_data

    return cleaned_dictionary


def status_reset(data_dict):
    # Iterate through each employee in the month
    for employee, employee_data in data_dict.items():
        # Check if the employee data has a Status key
        if 'Status' in employee_data:
            # Create a new status list to replace the old one
            new_status = []
            for status in employee_data['Status']:
                # Keep "WO" (Weekend Off) as is, change everything else to "NYD"
                if status == 'WO':
                    new_status.append('NYD')
                else:
                    new_status.append('NYD')

            # Update the Status list in the dictionary
            data_dict[employee]['Status'] = new_status

    return data_dict


def sunday_finder(data_dict):
    # Iterate through each employee in the month
    for employee, employee_data in data_dict.items():
        # Check if the employee data has both 'Days' and 'Status' keys
        if 'Days' in employee_data and 'Status' in employee_data:
            days_list = employee_data['Days']
            status_list = employee_data['Status']

            # Make sure the lists have the same length
            if len(days_list) == len(status_list):
                # Iterate through each day
                for i, day_string in enumerate(days_list):
                    # Add type checking before searching for "Sunday"
                    if isinstance(day_string, str) and 'Sunday' in day_string:
                        # Update the corresponding status to "WO"
                        status_list[i] = "WO"
                    # Enhanced debug print with month and employee info
                    elif not isinstance(day_string, str):
                        print(f"  Value: {day_string}")
                        print(f"  Type: {type(day_string)}")
                        print(f"  Full days list: {days_list[:5]}... (showing first 5 items)")

                # Update the Status list in the dictionary
                data_dict[employee]['Status'] = status_list

    return data_dict


###################################################################################################################
def get_day(date_str):
    from datetime import datetime
    return datetime.strptime(date_str, '%d %B %Y').strftime('%A')


holiday_dictionary = {
    2024: {
        "New Year": f"1 January 2024, {get_day('1 January 2024')}",
        "Republic Day": f"26 January 2024, {get_day('26 January 2024')}",
        "Holi": f"25 March 2024, {get_day('25 March 2024')}",
        "Ramzan Eid": f"9 April 2024, {get_day('9 April 2024')}",
        "Gudi Padwa": f"9 April 2024, {get_day('9 April 2024')}",
        "Labour Day": f"1 May 2024, {get_day('1 May 2024')}",
        "Independence Day": f"15 August 2024, {get_day('15 August 2024')}",
        "Raksha Bandhan": f"19 August 2024, {get_day('19 August 2024')}",
        "Ganesh Chaturthi": f"7 September 2024, {get_day('7 September 2024')}",
        "Gandhi Jayanti": f"2 October 2024, {get_day('2 October 2024')}",
        "Dusshera": f"12 October 2024, {get_day('12 October 2024')}",
        "Diwali": f"1 November 2024, {get_day('1 November 2024')}",
        "Diwali (Second Day)": f"2 November 2024, {get_day('2 November 2024')}",
        "Christmas": f"25 December 2024, {get_day('25 December 2024')}"
    },
    2025: {
        "New Year": f"1 January 2025, {get_day('1 January 2025')}",
        "Republic Day": f"26 January 2025, {get_day('26 January 2025')}",
        "office Picnic": f"1 February 2025, {get_day('1 February 2025')}",
        "Holi": f"14 March 2025, {get_day('14 March 2025')}",
        "Ramzan": f"31 March 2025, {get_day('31 March 2025')}",
        "Labour Day / Maharashtra Diwas": f"1 May 2025, {get_day('1 May 2025')}",
        "Raksha Bandhan": f"9 August 2025, {get_day('9 August 2025')}",
        "Independence Day": f"15 August 2025, {get_day('15 August 2025')}",
        "Ganesh Chaturthi": f"27 August 2025, {get_day('27 August 2025')}",
        "Gandhi Jayanti / Dussehra": f"2 October 2025, {get_day('2 October 2025')}",
        "Diwali": f"21 October 2025, {get_day('21 October 2025')}",
        "Bhai Duj": f"23 October 2025, {get_day('23 October 2025')}",
        "Christmas": f"25 December 2025, {get_day('25 December 2025')}"
    }
}


def daily_working_hours_calculation_bulk(employee_dict):
    from datetime import datetime, timedelta

    for employee, employee_data in employee_dict.items():
        in_times = employee_data['InTime']
        out_times = employee_data['OutTime']
        status_list = employee_data['Status']
        daily_working_hours = []

        for i, (in_time, out_time) in enumerate(zip(in_times, out_times)):
            if in_time != 'NaT' and out_time != 'NaT':
                # Function to parse time strings
                def parse_time(time_str):
                    if ' ' in time_str:  # Check if it's a full datetime string
                        time_str = time_str.split(' ')[1]  # Get the time part
                    return datetime.strptime(time_str, '%H:%M') if len(time_str) == 5 else datetime.strptime(time_str,
                                                                                                             '%H:%M:%S')

                try:
                    in_time_dt = parse_time(in_time)
                    out_time_dt = parse_time(out_time)

                    # Calculate the difference in time
                    working_duration = out_time_dt - in_time_dt

                    # Handle cases where the out time is on the next day
                    if working_duration < timedelta(0):
                        working_duration += timedelta(days=1)

                    # Convert duration to HH:MM format
                    hours, remainder = divmod(working_duration.total_seconds(), 3600)
                    minutes = remainder // 60
                    working_hours = f"{int(hours):02}:{int(minutes):02}"
                    daily_working_hours.append(working_hours)

                    # Condition 1: If status is "NYD", update it to "P"
                    if status_list[i] == "NYD":
                        status_list[i] = "P"

                    # Condition 2: If status is "WO", update it to "WOP"
                    elif status_list[i] == "WO":
                        status_list[i] = "WOP"

                except ValueError:
                    daily_working_hours.append("NaT")  # Append "NaT" if parsing fails
            else:
                daily_working_hours.append("NaT")  # Append "NaT" if in_time or out_time is 'NaT'

        # Update the employee's dailyWorkingHours and Status in the JSON
        employee_dict[employee]['dailyWorkingHours'] = daily_working_hours
        employee_dict[employee]['Status'] = status_list

    return employee_dict


def fixed_holidays(data_dict, holiday_dictionary):
    """
    Update attendance records based on fixed holidays for the new data structure.

    Rules:
    - If a day falls on a holiday and status is "NYD", change to "HO"
    - If a day falls on a holiday and status is "WO", keep as "WO" (already weekend)
    - If a day falls on a holiday and status is "P", change to "HOP" (present on holiday)

    Args:
        data_dict: Dictionary containing attendance data by employee
        holiday_dictionary: Dictionary with holiday dates by year

    Returns:
        Updated attendance dictionary
    """
    # Create a lookup dictionary for easy holiday checking
    holiday_lookup = {}

    # Convert holiday dates to lookup format (YYYY-MM-DD)
    for year, holidays in holiday_dictionary.items():
        for holiday_name, date_str in holidays.items():
            # Parse the date string (e.g., '1 January 2024, Monday')
            date_parts = date_str.split(',')[0].strip().split()
            day = int(date_parts[0])
            month_name = date_parts[1]
            year = int(date_parts[2])

            # Convert month name to number
            month_mapping = {
                'January': 1, 'February': 2, 'March': 3, 'April': 4,
                'May': 5, 'June': 6, 'July': 7, 'August': 8,
                'September': 9, 'October': 10, 'November': 11, 'December': 12
            }
            month = month_mapping[month_name]

            # Create date key in format YYYY-MM-DD
            date_key = f"{year}-{month:02d}-{day:02d}"
            holiday_lookup[date_key] = holiday_name

    # Process each employee in the attendance dictionary
    for employee_name, emp_data in data_dict.items():
        if isinstance(emp_data, dict) and 'Status' in emp_data and 'Days' in emp_data:
            # Get the status list and days list
            status_list = emp_data['Status']
            days_list = emp_data['Days']

            # Process each day's status
            for i, (status, day_str) in enumerate(zip(status_list, days_list)):
                # Parse the date from the day string (e.g., '15 August 2024, Thursday')
                if ',' in day_str:
                    date_parts = day_str.split(',')[0].strip().split()
                    if len(date_parts) >= 3:
                        day = int(date_parts[0])
                        month_name = date_parts[1]
                        year = int(date_parts[2])

                        # Convert month name to number
                        month_mapping = {
                            'January': 1, 'February': 2, 'March': 3, 'April': 4,
                            'May': 5, 'June': 6, 'July': 7, 'August': 8,
                            'September': 9, 'October': 10, 'November': 11, 'December': 12
                        }
                        month = month_mapping[month_name]

                        date_key = f"{year}-{month:02d}-{day:02d}"

                        # Check if this date is a holiday
                        if date_key in holiday_lookup:
                            # Apply rules based on current status
                            if status == "NYD":
                                data_dict[employee_name]['Status'][i] = "HO"
                            elif status == "P":
                                data_dict[employee_name]['Status'][i] = "HOP"
                            # If status is "WO", keep it as is (already weekend)

    return data_dict


def absent_days(employee_dict):
    for employee, records in employee_dict.items():
        for i in range(len(records['Status'])):
            if (
                    records['InTime'][i] == 'NaT' and
                    records['OutTime'][i] == 'NaT' and
                    records['Status'][i] == 'NYD'
            ):
                records['Status'][i] = 'A'  # Mark as Absent

    return employee_dict


def missing_punch(attendance_data):
    """
    Analyzes employee attendance data to identify missing punches and updates the data with average InTime and OutTime.

    Args:
        attendance_data (dict): Dictionary containing employee attendance data

    Returns:
        tuple: (updated_attendance_data, missing_punch_insights)
    """
    missing_punch_insights = {}

    for employee, data in attendance_data.items():
        employee_insights = []
        average_in_time = data['averageInTime']
        average_out_time = data['averageOutTime']

        for i in range(len(data['Days'])):
            day = data['Days'][i]
            status = data['Status'][i]
            in_time = data['InTime'][i]
            out_time = data['OutTime'][i]

            # Check if the day has NYD status or potential missing punches
            if status == 'NYD' or (in_time != 'NaT' and out_time == 'NaT'):
                # Handle case where employee forgot to punch out
                if in_time != 'NaT' and out_time == 'NaT':
                    in_time_hour = int(in_time.split(':')[0])

                    # If punch time is before 11:00, it's likely a proper in-time with missing out-time
                    if in_time_hour < 11:
                        employee_insights.append({
                            'day': day,
                            'issue': 'Missing punch-out',
                            'current_status': status,
                            'recommendation': 'Update OutTime'
                        })
                        # Update OutTime with averageOutTime
                        data['OutTime'][i] = average_out_time
                    # If punch time is 11:00 or after, it's likely a missing in-time but proper out-time
                    else:
                        employee_insights.append({
                            'day': day,
                            'issue': 'Missing punch-in, OutTime recorded as InTime',
                            'current_status': status,
                            'recommendation': 'Move InTime to OutTime, set InTime to NaT, update status to P'
                        })
                        # Move InTime to OutTime, set InTime to NaT
                        data['OutTime'][i] = in_time
                        data['InTime'][i] = 'NaT'
                        data['Status'][i] = 'P'

                # Handle NYD status with any punch recorded
                elif status == 'NYD':
                    if in_time != 'NaT':
                        in_time_hour = int(in_time.split(':')[0])

                        # If punch time is before 11:00, it's likely a proper in-time with missing out-time
                        if in_time_hour < 11:
                            employee_insights.append({
                                'day': day,
                                'issue': 'Missing punch-out',
                                'current_status': status,
                                'recommendation': 'Update OutTime, change status to P'
                            })
                            # Update OutTime with averageOutTime
                            data['OutTime'][i] = average_out_time
                            data['Status'][i] = 'P'
                        # If punch time is 11:00 or after, it's likely a missing in-time but proper out-time
                        else:
                            employee_insights.append({
                                'day': day,
                                'issue': 'Missing punch-in, OutTime recorded as InTime',
                                'current_status': status,
                                'recommendation': 'Move InTime to OutTime, set InTime to NaT, update status to P'
                            })
                            # Move InTime to OutTime, set InTime to NaT
                            data['OutTime'][i] = in_time
                            data['InTime'][i] = 'NaT'
                            data['Status'][i] = 'P'

        # Only add employee to insights if there are issues to report
        if employee_insights:
            missing_punch_insights[employee] = employee_insights

    return attendance_data, missing_punch_insights


def calculate_daily_working_hours(employee_dict):
    from datetime import datetime, timedelta
    # Function to parse time in HH:MM format
    def parse_time(time_str):
        return datetime.strptime(time_str, "%H:%M") if time_str != 'NaT' else None

    # Function to format timedelta to HH:MM
    def format_timedelta(td):
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}"

    # Function to calculate average time from a list of datetime objects
    def calculate_average_time(time_list):
        total_seconds = sum([(t.hour * 3600 + t.minute * 60) for t in time_list])
        average_seconds = total_seconds // len(time_list)
        hours, remainder = divmod(average_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}"

    for employee, details in employee_dict.items():
        in_times = details['InTime']
        out_times = details['OutTime']
        working_hours = []
        total_working_time = timedelta()
        valid_in_times = []
        valid_out_times = []

        for in_time, out_time in zip(in_times, out_times):
            if in_time != 'NaT' and out_time != 'NaT':
                in_time_parsed = parse_time(in_time)
                out_time_parsed = parse_time(out_time)

                # Handle the case where the employee worked past midnight
                if out_time_parsed < in_time_parsed:
                    out_time_parsed += timedelta(days=1)

                duration = out_time_parsed - in_time_parsed
                working_hours.append(format_timedelta(duration))
                total_working_time += duration
                valid_in_times.append(in_time_parsed)
                valid_out_times.append(out_time_parsed)
            else:
                working_hours.append('NaT')

        details['dailyWorkingHours'] = working_hours

        # Calculate average working time
        total_days = len([time for time in working_hours if time != 'NaT'])
        if total_days > 0:
            average_working_time = total_working_time / total_days
            details['averageWorkingHour'] = format_timedelta(average_working_time)
            details['averageInTime'] = calculate_average_time(valid_in_times)
            details['averageOutTime'] = calculate_average_time(valid_out_times)
        else:
            details['averageWorkingHour'] = '00:00'
            details['averageInTime'] = '00:00'
            details['averageOutTime'] = '00:00'

    return employee_dict


def half_day(employee_dict):
    for employee, details in employee_dict.items():
        half_day_map = [0] * len(details['Days'])

        for i in range(len(details['Days'])):
            daily_hours = details['dailyWorkingHours'][i]
            status = details['Status'][i]

            if daily_hours != 'NaT':
                hours, minutes = map(int, daily_hours.split(':'))
                total_minutes = hours * 60 + minutes

                if total_minutes < 420:  # less than 7 hours
                    half_day_map[i] = 1
                    if status == 'P':
                        details['Status'][i] = 'P1/2'
                    elif status == 'WO':
                        details['Status'][i] = 'WOP1/2'
                    elif status == 'HO':
                        details['Status'][i] = 'HOP1/2'

        details['halfDayMap'] = half_day_map
        details['halfDayTotal'] = half_day_map.count(1)

    return employee_dict


def recalibrator(employee_dict):
    from datetime import datetime, timedelta

    def parse_time(time_str):
        return datetime.strptime(time_str, "%H:%M") if time_str != 'NaT' else None

    def format_timedelta(td):
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}"

    for employee, details in employee_dict.items():
        for i in range(len(details['Days'])):
            status = details['Status'][i]
            in_time = details['InTime'][i]
            out_time = details['OutTime'][i]

            if status == 'NYD' and in_time != 'NaT' and out_time != 'NaT':
                in_time_parsed = parse_time(in_time)
                out_time_parsed = parse_time(out_time)

                # Handle the case where the employee worked past midnight
                if out_time_parsed < in_time_parsed:
                    out_time_parsed += timedelta(days=1)

                duration = out_time_parsed - in_time_parsed
                details['dailyWorkingHours'][i] = format_timedelta(duration)
                details['Status'][i] = 'P'

    return employee_dict


def calculate_latemark(employee_dict):
    for employee, data in employee_dict.items():
        in_time_list = data['InTime']  # Get the InTime list
        late_mark = []  # Initialize the lateMark list
        late_count = 0  # Counter for late marks

        # Initialize lateMarkAbsentee if not present
        if 'lateMarkAbsentee' not in data:
            data['lateMarkAbsentee'] = 0.0

        for in_time in in_time_list:
            if in_time and in_time != 'NaT':  # Check if the InTime is valid
                try:
                    hours, minutes = map(int, in_time.split(':'))
                    if (hours == 10 and minutes > 0) or (10 < hours < 12) or (hours == 12 and minutes == 0):
                        late_mark.append(1)  # Mark as late
                        late_count += 1
                    else:
                        late_mark.append(0)  # Not late
                except ValueError:
                    late_mark.append(0)  # Handle unexpected formats
            else:
                late_mark.append(0)  # Handle NaT cases

        # Calculate absentee days due to late marks
        data['lateMarkAbsentee'] += (late_count // 3) * 0.5

        # Store the lateMark in the employee's data
        data['lateMark'] = late_mark
        data['lateMarkCount'] = data['lateMark'].count(1)

    return employee_dict


def early_leave(employee_dict, expected_work_hours=9):
    for employee, data in employee_dict.items():
        status = data.get('Status', [])
        in_time = data.get('InTime', [])
        out_time = data.get('OutTime', [])

        early_leave_map = []  # List to store early leave mappings
        early_leave_time = []  # List to store early leave times in HH:MM format
        total_early_leave = 0  # Counter for total early leaves
        total_incomplete_minutes = 0  # Counter for total early leave time in minutes

        for i in range(len(status)):
            if status[i] == 'P':  # Check if status is P
                if in_time[i] != 'NaT' and out_time[i] != 'NaT':
                    # Parse InTime and OutTime
                    in_h, in_m = map(int, in_time[i].split(':'))
                    out_h, out_m = map(int, out_time[i].split(':'))

                    # Adjust for midnight crossing
                    if out_h < in_h or (out_h == in_h and out_m < in_m):
                        out_h += 24

                    # Calculate total working hours
                    total_working_minutes = (out_h * 60 + out_m) - (in_h * 60 + in_m)
                    total_working_hours = total_working_minutes / 60  # Convert to hours

                    # Calculate early leave time
                    expected_working_minutes = expected_work_hours * 60
                    if total_working_minutes < expected_working_minutes:
                        early_leave_map.append(1)  # Mark as early leave
                        total_early_leave += 1  # Increment early leave count

                        # Calculate early leave time in minutes
                        early_minutes = expected_working_minutes - total_working_minutes
                        total_incomplete_minutes += early_minutes

                        # Convert early leave time to HH:MM format
                        early_h = early_minutes // 60
                        early_m = early_minutes % 60
                        early_leave_time.append(f"{early_h:02}:{early_m:02}")
                    else:
                        early_leave_map.append(0)  # Did not leave early
                        early_leave_time.append("00:00")
                else:
                    early_leave_map.append(0)  # Handle NaT cases as not early leave
                    early_leave_time.append("00:00")
            else:
                early_leave_map.append(0)  # Ignore other statuses
                early_leave_time.append("00:00")

        # Convert total incomplete minutes to HH:MM format
        incomplete_h = total_incomplete_minutes // 60
        incomplete_m = total_incomplete_minutes % 60
        incomplete_hours = f"{incomplete_h:02}:{incomplete_m:02}"

        # Update the employee's dictionary with new data
        employee_dict[employee]['earlyLeaveMap'] = early_leave_map
        employee_dict[employee]['earlyLeaveTime'] = early_leave_time
        employee_dict[employee]['totalEarlyLeave'] = total_early_leave
        employee_dict[employee]['incompleteHours'] = incomplete_hours

    return employee_dict


def nonworking_days_compoff(employee_dict):
    from datetime import datetime, timedelta

    def parse_time(time_str):
        return datetime.strptime(time_str, "%H:%M") if time_str != 'NaT' else None

    for employee, details in employee_dict.items():
        comp_off = 0  # Initialize the compOff counter

        for i in range(len(details['Days'])):
            status = details['Status'][i]
            in_time = details['InTime'][i]
            out_time = details['OutTime'][i]

            # Check for working on weekends (WO)
            if status == 'WOP' and in_time != 'NaT' and out_time != 'NaT':
                comp_off += 1  # Increment compOff for each working weekend

            # Check for working on holidays (HO)
            if status == 'HOP' and in_time != 'NaT' and out_time != 'NaT':
                comp_off += 1  # Increment compOff for each working holiday

        # Update the employee's dictionary with the compOff value
        employee_dict[employee]['compOff'] = comp_off

    return employee_dict


def overtime(employee_dict, expected_work_hours=9):
    expected_work_minutes = expected_work_hours * 60  # Convert expected work hours to minutes
    from datetime import datetime, timedelta

    def format_timedelta(td):
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}"

    for employee, details in employee_dict.items():
        over_time = []
        total_actual_overtime_minutes = 0
        total_payable_overtime_minutes = 0

        working_hours = details.get('dailyWorkingHours', [])
        statuses = details.get('Status', [])

        for i in range(len(working_hours)):
            daily_hours = working_hours[i]
            status = statuses[i] if i < len(statuses) else None

            if daily_hours != 'NaT':
                hours, minutes = map(int, daily_hours.split(':'))
                total_minutes = hours * 60 + minutes

                # Check for WOP/WOS special case
                if status in ['WOP', 'WOS', 'WOP1/2']:
                    overtime_td = timedelta(minutes=total_minutes)
                    over_time.append(format_timedelta(overtime_td))
                    total_actual_overtime_minutes += total_minutes
                    total_payable_overtime_minutes += total_minutes
                elif total_minutes > expected_work_minutes:
                    overtime_minutes = total_minutes - expected_work_minutes
                    overtime_td = timedelta(minutes=overtime_minutes)
                    over_time.append(format_timedelta(overtime_td))
                    total_actual_overtime_minutes += overtime_minutes
                    if overtime_minutes > 60:
                        total_payable_overtime_minutes += overtime_minutes
                else:
                    over_time.append("00:00")
            else:
                over_time.append("00:00")

        details['overTime'] = over_time
        details['actualOverTime'] = format_timedelta(timedelta(minutes=total_actual_overtime_minutes))
        details['payableOverTime'] = format_timedelta(timedelta(minutes=total_payable_overtime_minutes))

    return employee_dict


def saturday_compoff(employee_dict):
    """
    Function to calculate CompOffTotal and update employee's Saturday statuses based on attendance rules.

    Args:
        employee_dict (dict): Dictionary containing employee data.

    Returns:
        dict: Updated employee dictionary with CompOffTotal and modified Saturday statuses.
    """
    for employee, data in employee_dict.items():
        # Initialize variables
        comp_off_total = 0
        status_list = data['Status']  # List of statuses for each day of the month
        days_list = data['Days']  # List of corresponding days

        absent_saturdays = 0
        working_saturdays = 0  # Saturdays excluding holidays

        # Iterate over days to evaluate Saturdays
        for i in range(len(days_list)):
            if 'Saturday' in days_list[i]:  # Check if the day is a Saturday
                if status_list[i] == 'HO':
                    continue  # Skip public holidays
                working_saturdays += 1
                if status_list[i] == 'A':
                    absent_saturdays += 1
                if status_list[i] == 'P1/2':
                    absent_saturdays += 0.5

        # Add 1 to CompOffTotal only if all working Saturdays are attended
        if absent_saturdays == 0 and working_saturdays > 0:
            comp_off_total += 1

        if absent_saturdays == 0.5 and working_saturdays > 0:
            comp_off_total += 0.5

        # Handle absences based on the number of missed Saturdays
        if absent_saturdays == 1:
            # Change the status of the missed Saturday to "WOS"
            for i in range(len(status_list)):
                if 'Saturday' in days_list[i] and status_list[i] == 'A':
                    status_list[i] = 'WOS'
                    break
        elif absent_saturdays == 2:
            # Change one "A" to "WOS" and keep the other as "A"
            wos_found = False
            for i in range(len(status_list)):
                if 'Saturday' in days_list[i]:
                    if status_list[i] == 'A' and not wos_found:
                        status_list[i] = 'WOS'
                        wos_found = True
                    elif status_list[i] == 'A':
                        status_list[i] = 'A'  # Keep as absent
        elif absent_saturdays >= 3:
            # Change one "A" to "WOS" and leave the rest as "A"
            wos_found = False
            for i in range(len(status_list)):
                if 'Saturday' in days_list[i]:
                    if status_list[i] == 'A' and not wos_found:
                        status_list[i] = 'WOS'
                        wos_found = True
                    elif status_list[i] == 'A':
                        status_list[i] = 'A'  # Keep as absent

        # Update employee data
        data['compOff'] = data.get('compOff', 0) + comp_off_total
        data['Status'] = status_list

    return employee_dict


def calculate_metric(employee_dict):
    for employee, data in employee_dict.items():
        # Initialize the reportMetric dictionary
        report_metric = {
            "CalenderDays": 0,
            "OfficeWorkingDays": 0,
            "EmployeeTotalWorkingDay": 0,
            "PublicHolidays": 0,
            "EmployeeAverageWorkingHours": "00:00",
            "EmployeeTotalWorkingHours": "00:00",
            "EmployeeActualAbsentee": 0,
            "TotalHolidays": 0
        }
        # Calculation 0: CalenderDays
        total_days = len(data['Days'])
        report_metric["CalenderDays"] = total_days

        # Calculation 1: OfficeWorkingDays
        total_days = len(data['Days'])
        holidays = data['Status'].count('HO')
        sundays = sum(1 for day in data['Days'] if 'Sunday' in day)
        report_metric["OfficeWorkingDays"] = total_days - holidays - sundays - 1

        # Calculation 2: EmployeeTotalWorkingDay
        total_working_days = 0
        for status in data['Status']:
            if status == 'P':
                total_working_days += 1
            elif status == 'P1/2':
                total_working_days += 0.5
            elif status in ['WOP']:
                total_working_days += 1
            elif status in ['WOP1/2']:
                total_working_days += 0.5
        report_metric["EmployeeTotalWorkingDay"] = total_working_days

        # Calculation 3: PublicHolidays
        report_metric["PublicHolidays"] = holidays

        # Calculation 4: EmployeeAverageWorkingHours
        total_working_hours = timedelta()
        valid_days = 0
        for hours in data['dailyWorkingHours']:
            if hours != 'NaT':
                h, m = map(int, hours.split(':'))
                total_working_hours += timedelta(hours=h, minutes=m)
                valid_days += 1

        if valid_days > 0:
            avg_working_hours = total_working_hours / valid_days
            avg_hours, rem = divmod(avg_working_hours.total_seconds(), 3600)
            avg_minutes = rem // 60
            report_metric["EmployeeAverageWorkingHours"] = f"{int(avg_hours):02}:{int(avg_minutes):02}"

        # Calculation 5: EmployeeTotalWorkingHours
        total_hours, rem = divmod(total_working_hours.total_seconds(), 3600)
        total_minutes = rem // 60
        report_metric["EmployeeTotalWorkingHours"] = f"{int(total_hours):02}:{int(total_minutes):02}"

        # Calculation 6: EmployeeTotalAbsentee
        report_metric["EmployeeActualAbsentee"] = data['Status'].count('A')

        # Calculation 7: EmployeeTotalAbsentee including Late marks
        report_metric["EmployeeAbsenteeWithLateMark"] = data['Status'].count('A') + data['lateMarkAbsentee']

        # Calculation 8: Total Holidays (HO, WOS, WO)
        total_holidays = (
            data['Status'].count('HO') +
            data['Status'].count('WOS') +
            data['Status'].count('WO')
        )
        report_metric["TotalHolidays"] = total_holidays

        # Add reportMetric to the employee's data
        data['reportMetric'] = report_metric

    return employee_dict

def finalAdjustment(employee_dict):
    for employee, data in employee_dict.items():
        office_working_days = data['reportMetric']['OfficeWorkingDays']
        employee_total_working_day = data['reportMetric']['EmployeeTotalWorkingDay']

        # Check if EmployeeTotalWorkingDay is greater than OfficeWorkingDays
        if employee_total_working_day > office_working_days:
            # Calculate the difference
            difference = employee_total_working_day - office_working_days

            # Update CompOffTotal
            data['compOff'] += difference
            data['reportMetric']['EmployeeTotalWorkingDay'] -= difference

    return employee_dict


###################################################################################################################
def calculate_adherence_ratio(employee_dict):
    """
    Calculate the adherence ratio for each employee based on their late mark count and total office working days,
    and add it to their respective dictionaries.

    Parameters:
    employee_dict (dict): Dictionary containing employee attendance and work details.

    Returns:
    dict: Updated employee dictionary with adherence ratio key added inside each employee's dictionary.
    """
    for employee_name, employee_data in employee_dict.items():
        late_mark_count = employee_data.get('lateMarkCount', 0)
        total_employee_working_days = employee_data.get('reportMetric', {}).get('EmployeeTotalWorkingDay',
                                                                                1)  # Default to 1 to avoid division by zero

        # Ensure total_office_working_days is at least 1 to avoid division by zero
        if total_employee_working_days <= 0:
            total_employee_working_days = 1  # Set to 1 to avoid division by zero

        # Calculate adherence ratio
        adherence_ratio = late_mark_count / total_employee_working_days
        adherence_ratio = round(adherence_ratio, 2)
        employee_data['reportMetric'] = employee_data.get('reportMetric', {})
        employee_data['reportMetric']['adherenceRatio'] = adherence_ratio

    for employee_name, employee_data in employee_dict.items():
        adherence_ratio = employee_data['reportMetric'].get('adherenceRatio', 0)

        # New conditions based on your requirements
        if employee_data['lateMarkCount'] == 0:
            adherence_ratio_star = 5  # 5 stars if no late marks
        elif adherence_ratio < 0.1:  # Adjusted for very low adherence ratios
            adherence_ratio_star = 5  # 5 stars for very low adherence ratio
        elif 0.1 <= adherence_ratio < 0.2:
            adherence_ratio_star = 4
        elif 0.2 <= adherence_ratio < 0.4:
            adherence_ratio_star = 3
        elif 0.4 <= adherence_ratio < 0.6:
            adherence_ratio_star = 2
        elif 0.6 <= adherence_ratio < 0.8:
            adherence_ratio_star = 1
        elif 0.8 <= adherence_ratio <= 1.0:
            adherence_ratio_star = 0  # Invalid adherence ratio
        else:
            adherence_ratio_star = 0  # Invalid adherence ratio

        employee_data['reportMetric']['adherenceRatioStar'] = adherence_ratio_star

    return employee_dict


def calculate_work_deficit_ratio(employee_dict):
    def time_to_minutes(time_str):
        """Convert a time string in HH:MM format to total minutes."""
        if time_str == 'NaT' or time_str == '00:00' or time_str == '0':
            return 0
        hours, minutes = map(int, time_str.split(':'))
        return hours * 60 + minutes

    """
    Calculate the Work Deficit Ratio for each employee based on their working hours,
    and add it to their respective dictionaries.

    Parameters:
    employee_dict (dict): Dictionary containing employee attendance and work details.

    Returns:
    dict: Updated employee dictionary with workDeficitRatio and workDeficitRatioStar keys added inside each employee's dictionary.
    """
    for employee_name, employee_data in employee_dict.items():
        # Extracting the required values and converting them to minutes
        incomplete_working_hours = time_to_minutes(employee_data.get('incompleteHours', '00:00'))
        overtime_hours = time_to_minutes(employee_data.get('payableOverTime', '00:00'))
        total_working_hours = time_to_minutes(
            employee_data['reportMetric'].get('EmployeeTotalWorkingHours', '00:00'))  # Default to '00:00'

        # Calculate Work Deficit Ratio
        if total_working_hours > 0:  # Avoid division by zero
            work_deficit_ratio = (incomplete_working_hours - overtime_hours) / total_working_hours
        else:
            work_deficit_ratio = 0  # If total working hours is 0, set ratio to 0

        # Add the calculated ratio to the employee's data
        employee_data['reportMetric']['workDeficitRatio'] = round(work_deficit_ratio, 2)

        # Determine the star rating based on the work deficit ratio
        if work_deficit_ratio <= -0.05:
            employee_data['reportMetric']['workDeficitRatioStar'] = 5  # ⭐⭐⭐⭐⭐
        elif -0.05 < work_deficit_ratio <= -0.01:
            employee_data['reportMetric']['workDeficitRatioStar'] = 4  # ⭐⭐⭐⭐
        elif -0.01 < work_deficit_ratio <= 0.01:
            employee_data['reportMetric']['workDeficitRatioStar'] = 3  # ⭐⭐⭐
        elif 0.01 < work_deficit_ratio <= 0.05:
            employee_data['reportMetric']['workDeficitRatioStar'] = 2  # ⭐⭐
        elif 0.05 < work_deficit_ratio < 0.1:
            employee_data['reportMetric']['workDeficitRatioStar'] = 1  # ⭐
        else:  # work_deficit_ratio >= 0.1
            employee_data['reportMetric']['workDeficitRatioStar'] = 0  # 🚨 (0 Stars)

    return employee_dict


def calculate_adjusted_absentee_rate(employee_dict):
    """
    Calculate the Adjusted Absentee Rate for each employee based on their absenteeism and working days,
    and add it to their respective dictionaries.

    Parameters:
    employee_dict (dict): Dictionary containing employee attendance and work details.

    Returns:
    dict: Updated employee dictionary with adjustedAbsenteeRate and adjustedAbsenteeRateStar keys added inside each employee's dictionary.
    """
    for employee_name, employee_data in employee_dict.items():
        # Extracting the required values
        absentee_with_late_mark = employee_data['reportMetric'].get('EmployeeAbsenteeWithLateMark', 0)
        office_working_days = employee_data['reportMetric'].get('OfficeWorkingDays',
                                                                1)  # Default to 1 to avoid division by zero

        # Calculate Adjusted Absentee Rate
        adjusted_absentee_rate = (absentee_with_late_mark / office_working_days) * 100 if office_working_days > 0 else 0
        adjusted_absentee_rate = round(adjusted_absentee_rate, 0)
        # Add the calculated rate to the employee's data
        employee_data['reportMetric']['adjustedAbsenteeRate'] = adjusted_absentee_rate

        # Determine the star rating based on the adjusted absentee rate
        if adjusted_absentee_rate <= 20:
            employee_data['reportMetric']['adjustedAbsenteeRateStar'] = 5  # ⭐⭐⭐⭐⭐
        elif 21 <= adjusted_absentee_rate <= 40:
            employee_data['reportMetric']['adjustedAbsenteeRateStar'] = 4  # ⭐⭐⭐⭐
        elif 41 <= adjusted_absentee_rate <= 60:
            employee_data['reportMetric']['adjustedAbsenteeRateStar'] = 3  # ⭐⭐⭐
        elif 61 <= adjusted_absentee_rate <= 80:
            employee_data['reportMetric']['adjustedAbsenteeRateStar'] = 2  # ⭐⭐
        else:  # Above 80%
            employee_data['reportMetric']['adjustedAbsenteeRateStar'] = 1  # ⭐

    return employee_dict

###################################################################################

def process_missing_data(insights_dict):
    import pandas as pd
    import re
    from datetime import datetime
    """
    Process missing attendance data and create a dataframe for reporting.

    Parameters:
    insights_dict (dict): Dictionary containing employee attendance issues

    Returns:
    pandas.DataFrame: DataFrame containing all missing data with recommendations
    """
    # List to store all records
    records = []

    # Extract month and year from the first record to use for filtering
    first_employee = list(insights_dict.keys())[0]
    first_record = insights_dict[first_employee][0]['day']
    month_year_match = re.search(r'(\w+)\s+(\d{4})', first_record)

    if month_year_match:
        month, year = month_year_match.groups()
    else:
        # Default to January 2025 if pattern not found
        month, year = "January", "2025"

    # Process each employee's records
    for employee_name, issues in insights_dict.items():
        for issue in issues:
            # Extract date from the day string
            day_str = issue['day']
            date_match = re.search(r'(\d+)\s+(\w+)\s+(\d{4})', day_str)

            if date_match:
                day, month, year = date_match.groups()
                try:
                    # Convert to datetime for sorting
                    date_obj = datetime.strptime(f"{day} {month} {year}", "%d %B %Y")
                    date_formatted = date_obj.strftime("%Y-%m-%d")
                except ValueError:
                    # If date parsing fails, use a placeholder
                    date_formatted = f"{year}-{month}-{day}"

                # Create a record for this issue
                record = {
                    'Employee Name': employee_name,
                    'Date': date_formatted,
                    'Day': day_str.split(',')[1].strip() if ',' in day_str else '',
                    'Issue': issue['issue'],
                    'Current Status': issue['current_status'],
                    'Recommendation': issue['recommendation']
                }
                records.append(record)

    # Create DataFrame from records
    df = pd.DataFrame(records)

    # Sort by employee name and date
    if not df.empty:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.sort_values(['Employee Name', 'Date'])

    # Reset index for clean output
    df = df.reset_index(drop=True)

    return df


def absentee_map(employee_dict):
    # Define statuses that are considered as absentee
    absentee_statuses = ['A']

    # Loop through each employee's data
    for employee, data in employee_dict.items():
        # Extract the status list from the employee's data
        status_list = data['Status']

        # Create the absentee map by checking if each status is in the absentee_statuses list
        absentee_map = [1 if status in absentee_statuses else 0 for status in status_list]

        # Add the absentee map to the employee's data
        data['absenteeMap'] = absentee_map

    return employee_dict
