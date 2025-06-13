import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, session
import os
import csv
import ast
import plotly.io
from plotly.io import to_html
from werkzeug.utils import secure_filename


from functions.hrone_functions import (process_employee_hroneData,
                                       dict_cleaning_hrone,
                                       update_weekdays_hrone,
                                       matching_mechanism)

from functions.dashboard_function_new import *
from functions.biometric_function_new import *

app = Flask(__name__)
app.secret_key = '123'  # Set a secret key for session management

# Use os.path.join consistently for all paths
UPLOAD_FOLDER = os.path.join('static', 'resources', 'uploads')
UPLOAD_FOLDER_BIOMETRIC = os.path.join('static', 'resources', 'uploads', 'BIOMETRIC_DATA')
UPLOAD_FOLDER_HRONE = os.path.join('static', 'resources', 'uploads', 'HRONE_DATA')

# Set Flask config values
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['UPLOAD_FOLDER_BIOMETRIC'] = UPLOAD_FOLDER_BIOMETRIC
app.config['UPLOAD_FOLDER_HRONE'] = UPLOAD_FOLDER_HRONE
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx', 'csv'}

CREDENTIALS_FILE = os.path.join('static', 'resources', 'user_credentials', 'login_credential.csv')

# Helper function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Function to validate user credentials
def validate_credentials(user_id, password):
    try:
        with open(CREDENTIALS_FILE, mode='r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['user_id'] == user_id and row['password'] == password:
                    return row['access'], row['name']
    except Exception as e:
        print(f"Error reading credentials file: {e}")
    return None, None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['POST'])
def login():
    user_id = request.form.get('user_id')
    password = request.form.get('password')

    access, name = validate_credentials(user_id, password)
    if access:
        session['user_id'] = user_id  # Store user ID in session
        session['access'] = access      # Store access level in session
        session['name'] = name          # Store username in session
        if access == 'admin':
            return redirect(url_for('admin'))
        elif access == 'user':
            return redirect(url_for('home'))
    else:
        flash('Invalid credentials. Please try again.', 'danger')
        return redirect(url_for('index'))


################################ Home ##################################
employee_dict = {}
insights = {}


@app.route('/home')
def home():
    load_saved_paths()
    global report_html

    global employee_dict
    global insights

    employee_dict = process_attendance_file(BIOMETRICPATH)
    employee_dict = date_cleaning(employee_dict)
    employee_dict = status_reset(employee_dict)
    employee_dict = sunday_finder(employee_dict)
    employee_dict = daily_working_hours_calculation_bulk(employee_dict)
    employee_dict = fixed_holidays(employee_dict, holiday_dictionary)
    employee_dict = absent_days(employee_dict)
    employee_dict = calculate_daily_working_hours(employee_dict)
    employee_dict, insights = missing_punch(employee_dict)
    employee_dict = recalibrator(employee_dict)
    employee_dict = half_day(employee_dict)
    employee_dict = calculate_latemark(employee_dict)
    employee_dict = early_leave(employee_dict)
    employee_dict = nonworking_days_compoff(employee_dict)
    employee_dict = overtime(employee_dict)
    employee_dict = saturday_compoff(employee_dict)
    employee_dict = calculate_metric(employee_dict)
    employee_dict = finalAdjustment(employee_dict)
    employee_dict = absentee_map(employee_dict)

    ################### Ratios Calculation ##################

    employee_dict = calculate_adherence_ratio(employee_dict)
    employee_dict = calculate_work_deficit_ratio(employee_dict)
    employee_dict = calculate_adjusted_absentee_rate(employee_dict)


    # First, extract the data you need from each employee record
    report_data = {}
    for employee, data in employee_dict.items():
        employee_data = {
            'employeeId' : data['EmployeeID'],
            'OfficeWorkingDays': data['reportMetric']['OfficeWorkingDays'],
            'PublicHolidays': data['reportMetric']['PublicHolidays'],
            'EmployeeTotalWorkingDay': data['reportMetric']['EmployeeTotalWorkingDay'],
            'EmployeeTotalWorkingHours': data['reportMetric']['EmployeeTotalWorkingHours'],
            'averageWorkingHour': data['averageWorkingHour'],
            'incompleteHours': data['incompleteHours'],
            'actualOverTime': data['actualOverTime'],
            'payableOverTime': data['payableOverTime'],
            'halfDayTotal': data['halfDayTotal'],
            'lateMarkCount': data['lateMarkCount'],
            'totalEarlyLeave': data['totalEarlyLeave'],

            'compOff': data['compOff'],
            'EmployeeActualAbsentee': data['reportMetric']['EmployeeActualAbsentee'],
            'EmployeeAbsenteeWithLateMark': data['reportMetric']['EmployeeAbsenteeWithLateMark'],

            # Main level fields

            'averageInTime': data['averageInTime'],
            'averageOutTime': data['averageOutTime'],

        }
        report_data[employee] = employee_data

    # Create DataFrame from the flattened dictionary
    reportDataframe = pd.DataFrame.from_dict(report_data, orient='index')
    reportDataframe.reset_index(inplace=True)
    reportDataframe.rename(columns={'index': 'Employee'}, inplace=True)

    # Rename columns as per your requirements
    reportDataframe.rename(columns={
        'employeeId' : 'Biometric Id',
        'OfficeWorkingDays': 'Office Working',
        'EmployeeTotalWorkingDay': 'Employee Total Present',
        'PublicHolidays': 'Public Holiday',
        'EmployeeTotalWorkingHours': 'Employee Total Working Hours',
        'EmployeeActualAbsentee': 'Physical Absentee',
        'lateMarkAbsentee': 'Late Mark Absentee',
        'lateMarkCount': 'Total Late Mark',
        'EmployeeAbsenteeWithLateMark': 'Employee Total Absentee',
        'compOff': 'Compensatory Off',
        'actualOverTime': 'Over Time',
        'payableOverTime': 'Payable Over Time',
        'incompleteHours': 'Incomplete Hours',
        'halfDayTotal': 'Total Half Days',
        'totalEarlyLeave': 'Total Early Leaves',
        'averageWorkingHour': 'Average Working Hours',
        'averageInTime': 'Average In Time',
        'averageOutTime': 'Average Out Time',
    }, inplace=True)

    # Convert DataFrame to HTML
    report_html = reportDataframe.to_html(index=False)

    # Retrieve the user's name from the session
    user_name = session.get('name', 'User ')  # Default to 'User ' if not found
    return render_template('home.html', user_name=user_name)


@app.route('/user_dashboard', methods=['GET', 'POST'])
def user_dashboard():
    global employee_dict  # Ensure you're using the global variable
    employee_names = list(employee_dict.keys())  # Extract employee names

    employee_dict_dashboard = {}
    total_work_hour_card = ""
    average_work_hour_card = ""
    actual_absantee_card = ""
    late_mark_card = ""
    selected_employee = ""
    target_gauge_html = ""
    daily_working_trend_line_html = ""
    status_donutChart_html = ""
    heatmap_metric_html = ""
    total_deduction_card = ""
    overtime_barchart_html = ""
    star_fig = ""

    # Handle POST request or default to first employee on GET request
    if request.method == 'POST':
        selected_employee = request.form.get('selected_employee')
    else:
        # Default to the first employee when the page loads (GET request)
        if employee_names:
            selected_employee = employee_names[0]

    # Process the selected employee data
    if selected_employee in employee_dict:
        employee_dict_dashboard = employee_dict[selected_employee]
        total_work_hour_card = total_working_hours(employee_dict_dashboard)
        average_work_hour_card = average_working_hours(employee_dict_dashboard)
        actual_absantee_card = acutal_absantees(employee_dict_dashboard)
        late_mark_card = late_marks_total(employee_dict_dashboard)
        total_deduction_card = total_deduction(employee_dict_dashboard)

        target_gauge = create_gauge_chart(employee_dict_dashboard)
        target_gauge_html = to_html(target_gauge, full_html=False)

        daily_working_trend_line = create_line_chart(employee_dict_dashboard)
        daily_working_trend_line_html = to_html(daily_working_trend_line, full_html=False)

        status_donutChart = create_donut_chart(employee_dict_dashboard)
        status_donutChart_html = to_html(status_donutChart, full_html=False)

        heatmap_metric = create_combined_barchart(employee_dict_dashboard)
        heatmap_metric_html = to_html(heatmap_metric, full_html=False)

        overtime_barchart = create_overtime_barchart(employee_dict_dashboard)
        overtime_barchart_html = to_html(overtime_barchart, full_html=False)

        star_fig = generate_star_rating_html(employee_dict_dashboard)

    return render_template('user_dashboard.html',
                           selected_employee=selected_employee,
                           employee_names=employee_names,
                           employee_dict=employee_dict,
                           employee_dict_dashboard=employee_dict_dashboard,
                           total_work_hour_card=total_work_hour_card,
                           average_work_hour_card=average_work_hour_card,
                           actual_absantee_card=actual_absantee_card,
                           late_mark_card=late_mark_card,
                           total_deduction_card=total_deduction_card,
                           target_gauge_html=target_gauge_html,
                           daily_working_trend_line_html=daily_working_trend_line_html,
                           status_donutChart_html=status_donutChart_html,
                           heatmap_metric_html=heatmap_metric_html,
                           overtime_barchart_html=overtime_barchart_html,
                           star_fig=star_fig)


@app.route('/user_report')
def user_report():

    missing_data = process_missing_data(insights)
    missing_data_html = missing_data.to_html(index=False)


    return render_template('user_report.html',
                           report_html=report_html,
                           missing_data_html = missing_data_html)

############################# ADMIN ################################
@app.route('/admin')
def admin():
    load_saved_paths()
    global report_html

    global employee_dict
    global insights

    employee_dict = process_attendance_file(BIOMETRICPATH)
    employee_dict = date_cleaning(employee_dict)
    employee_dict = status_reset(employee_dict)
    employee_dict = sunday_finder(employee_dict)
    employee_dict = daily_working_hours_calculation_bulk(employee_dict)
    employee_dict = fixed_holidays(employee_dict, holiday_dictionary)
    employee_dict = absent_days(employee_dict)
    employee_dict = calculate_daily_working_hours(employee_dict)
    employee_dict, insights = missing_punch(employee_dict)
    employee_dict = recalibrator(employee_dict)
    employee_dict = half_day(employee_dict)
    employee_dict = calculate_latemark(employee_dict)
    employee_dict = early_leave(employee_dict)
    employee_dict = nonworking_days_compoff(employee_dict)
    employee_dict = overtime(employee_dict)
    employee_dict = saturday_compoff(employee_dict)
    employee_dict = calculate_metric(employee_dict)
    employee_dict = finalAdjustment(employee_dict)
    employee_dict = absentee_map(employee_dict)
    ################### Ratios Calculation ##################

    employee_dict = calculate_adherence_ratio(employee_dict)
    employee_dict = calculate_work_deficit_ratio(employee_dict)
    employee_dict = calculate_adjusted_absentee_rate(employee_dict)


    # First, extract the data you need from each employee record
    report_data = {}
    for employee, data in employee_dict.items():
        employee_data = {
            'employeeId': data['EmployeeID'],
            'CalenderDays': data['reportMetric']['CalenderDays'],
            'TotalHolidays': data['reportMetric']['TotalHolidays'],
            # 'OfficeWorkingDays': data['reportMetric']['OfficeWorkingDays'],
            # 'PublicHolidays': data['reportMetric']['PublicHolidays'],
            'EmployeeTotalWorkingDay': data['reportMetric']['EmployeeTotalWorkingDay'],
            'EmployeeActualAbsentee': data['reportMetric']['EmployeeActualAbsentee'],
            'EmployeeAbsenteeWithLateMark': data['reportMetric']['EmployeeAbsenteeWithLateMark'],
            'EmployeeTotalWorkingHours': data['reportMetric']['EmployeeTotalWorkingHours'],
            'averageWorkingHour': data['averageWorkingHour'],
            'incompleteHours': data['incompleteHours'],
            'actualOverTime': data['actualOverTime'],
            'payableOverTime': data['payableOverTime'],
            'lateMarkCount': data['lateMarkCount'],
            'totalEarlyLeave': data['totalEarlyLeave'],
            'compOff': data['compOff'],

            # Main level fields
            'averageInTime': data['averageInTime'],
            'averageOutTime': data['averageOutTime'],

        }
        report_data[employee] = employee_data

    # Create DataFrame from the flattened dictionary
    reportDataframe = pd.DataFrame.from_dict(report_data, orient='index')
    reportDataframe.reset_index(inplace=True)
    reportDataframe.rename(columns={'index': 'Employee'}, inplace=True)

    # Rename columns as per your requirements
    reportDataframe.rename(columns={
        'employeeId': 'Biometric Id',
        'CalenderDays': 'Calender Days',
        'TotalHolidays': 'Total Holidays',
        # 'OfficeWorkingDays': 'Office Working',
        'EmployeeTotalWorkingDay': 'Employee Present',
        'EmployeeActualAbsentee': 'Physical Absentee',
        'EmployeeAbsenteeWithLateMark': 'Employee Total Absentee',
        # 'PublicHolidays': 'Public Holiday',
        'EmployeeTotalWorkingHours': 'Employee Total Working Hours',
        'lateMarkAbsentee': 'Late Mark Absentee',
        'lateMarkCount': 'Total Late Mark',
        'compOff': 'Compensatory Off',
        'actualOverTime': 'Over Time',
        'payableOverTime': 'Payable Over Time',
        'incompleteHours': 'Incomplete Hours',
        'totalEarlyLeave': 'Total Early Leaves',
        'averageWorkingHour': 'Average Working Hours',
        'averageInTime': 'Average In Time',
        'averageOutTime': 'Average Out Time',

    }, inplace=True)

    # Convert DataFrame to HTML
    report_html = reportDataframe.to_html(index=False)

    # Retrieve the user's name from the session
    user_name = session.get('name', 'User ')  # Default to 'User ' if not found
    return render_template('admin.html', user_name=user_name)


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    global BIOMETRICPATH
    global HRONEPATH

    # Define the path for saving file paths
    paths_file = os.path.join('static', 'paths.txt')

    if request.method == 'POST':
        biometric_file = request.files.get('biometric_file')
        uploaded_files = []

        # Create a dictionary to store paths
        file_paths = {}

        # Save the biometric file
        if biometric_file and allowed_file(biometric_file.filename):
            # Get original filename and secure it
            biometric_filename = secure_filename(biometric_file.filename)
            biometric_file_path = os.path.join(app.config['UPLOAD_FOLDER_BIOMETRIC'], biometric_filename)
            biometric_file.save(biometric_file_path)
            uploaded_files.append(biometric_filename)
            BIOMETRICPATH = biometric_file_path
            file_paths['bio_path'] = BIOMETRICPATH

        # Save the paths to a file
        if file_paths:
            with open(paths_file, 'w') as f:
                for key, value in file_paths.items():
                    f.write(f"{key}:{value}\n")

        return render_template('upload_success.html', files=uploaded_files)

    return render_template('uploads.html')


@app.route('/record', methods=['GET', 'POST'])
def record():
    if request.method == 'POST':
        # Handle file uploads and processing logic
        return "Processing files..."
    return render_template('record_uploading.html')



@app.route('/logout')
def logout():
    session.clear()  # Clear all session data
    flash('You have been logged out successfully.',  'success')
    return redirect(url_for('index'))


def load_saved_paths():
    global BIOMETRICPATH
    global HRONEPATH

    paths_file = os.path.join('static', 'paths.txt')

    if os.path.exists(paths_file):
        with open(paths_file, 'r') as f:
            for line in f:
                key, value = line.strip().split(':', 1)  # Split on first colon only
                if key == 'bio_path':
                    BIOMETRICPATH = value
                elif key == 'hrone_path':
                    HRONEPATH = value

if __name__ == '__main__':

    app.run(debug=True)

