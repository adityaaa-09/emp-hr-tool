import plotly.io
import plotly.express as px
from datetime import timedelta, datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots


def generate_employee_card(selected_employee):
    # Create the header for the employee name
    header_html = f"<h2 class='text-center font-weight-bold mb-4'>{selected_employee}</h2>"

def total_working_hours(employee_dict_dashboard):
    total_working_hour = employee_dict_dashboard.get('reportMetric', {}).get('EmployeeTotalWorkingHours', 'N/A')

    card =  f"""
        <div class="card m-2 card-total-working-hours">
            <div class="card-body">
                <h6>Total Working Hours</h6>
                <p>{total_working_hour}</p>
            </div>
        </div>
        """
    return card

def average_working_hours(employee_dict_dashboard):
    average_working_hour = employee_dict_dashboard.get('averageWorkingHour', 'N/A')

    card =  f"""
        <div class="card m-2 card-average-working-hours">
            <div class="card-body">
                <h6>Avg Working Hours</h6>
                <p>{average_working_hour}</p>
            </div>
        </div>
        """
    return card

def acutal_absantees(employee_dict_dashboard):
    actual_absantee = employee_dict_dashboard.get('reportMetric', {}).get('EmployeeActualAbsentee', 'N/A')

    card =  f"""
        <div class="card m-2 card-actual-absentees">
            <div class="card-body">
                <h6>Actual Absantees</h6>
                <p>{actual_absantee}</p>
            </div>
        </div>
        """
    return card

def late_marks_total(employee_dict_dashboard):
    late_mark = employee_dict_dashboard.get('lateMarkCount', 'N/A')

    card =  f"""
        <div class="card m-2 card-late-marks">
            <div class="card-body">
                <h6>Late Marks Deduction</h6>
                <p>{late_mark}</p>
            </div>
        </div>
        """
    return card

def total_deduction(employee_dict_dashboard):
    total_leave_deduction = employee_dict_dashboard.get('reportMetric', {}).get('EmployeeAbsenteeWithLateMark', 'N/A')

    card =  f"""
        <div class="card m-2 card-total-leave-deduction">
            <div class="card-body">
                <h6>Total Leave Deductions</h6>
                <p>{total_leave_deduction}</p>
            </div>
        </div>
        """
    return card

def create_gauge_chart(employee_dict_dashboard):
    office_working_days = employee_dict_dashboard['reportMetric']['OfficeWorkingDays']
    total_working_hours_str = employee_dict_dashboard['reportMetric']['EmployeeTotalWorkingHours']

    # Calculate expected working hours
    expected_working_hours = office_working_days * 9

    # Convert total working hours string to hours
    total_working_hours_parts = total_working_hours_str.split(':')
    total_working_hours = int(total_working_hours_parts[0]) + int(total_working_hours_parts[1]) / 60

    # Determine the color based on whether total working hours exceed expected working hours
    if total_working_hours >= expected_working_hours:
        bar_color = "#018749"
    else:
        bar_color = "#1B262C"

    # Create the gauge chart
    # Adjust the font size for delta and number
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=total_working_hours,
        number={'suffix': " Hr", 'font': {'size': 50}},  # Adjust the font size for the number
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': f"Working Hours", 'font': {'size': 20}},
        delta={'reference': expected_working_hours, 'increasing': {'color': "#018749"}, 'font': {'size': 20}},
        # Adjust the font size for delta
        gauge={
            'axis': {'range': [None, expected_working_hours * 1.2], 'tickwidth': 1, 'tickcolor': "black",
                     'tickmode': 'array', 'tickvals': list(range(0, int(expected_working_hours * 1.2) + 1, 25))},
            'bar': {'color': bar_color, 'thickness': 1},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "black",
            'steps': [
                {'range': [0, expected_working_hours * 0.5], 'color': 'white'},
                {'range': [expected_working_hours * 0.5, expected_working_hours], 'color': 'white'}],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.95,
                'value': expected_working_hours}}))

    fig.update_layout(
        paper_bgcolor="rgba(255, 255, 255, 0)",
        plot_bgcolor="rgba(255, 255, 255, 0)",
        font={'color': "black", 'family': "Arial"},
        margin=dict(l=50, r=50, t=80, b=80),  # Set margins to 0 to reduce background space
        width=500,  # Adjust the width as needed
        height=350,  # Adjust the height as needed
    )

    return fig

def create_line_chart(employee_dict_dashboard):
    employee_data = employee_dict_dashboard
    days = employee_data['Days']
    status = employee_data['Status']
    daily_working_hours = employee_data['dailyWorkingHours']

    # Filter out days with status 'HO' or 'WO'
    filtered_dates = []
    filtered_hours = []
    for day, stat, hours in zip(days, status, daily_working_hours):
        if stat not in ['HO', 'WO', 'WOS']:
            filtered_dates.append(day)
            filtered_hours.append(hours)

    # Convert 'NaT' to 0 and other time strings to hours
    filtered_hours = [0 if hours == 'NaT' else int(hours.split(':')[0]) + int(hours.split(':')[1]) / 60 for hours in
                      filtered_hours]

    # Convert string dates to datetime objects
    dates = [datetime.strptime(day.split(',')[0], '%d %B %Y') for day in filtered_dates]
    formatted_dates = [date.strftime('%d') for date in dates]

    # Extract the month from the first date for the title
    month = dates[0].strftime('%B') if dates else ""

    # Create the line chart
    fig = px.line(x=formatted_dates, y=filtered_hours, labels={'x': 'Date', 'y': 'Working Hours'},
                  title=f'Daily Working Hours for {month}')
    fig.update_traces(mode='lines+markers', line=dict(color='#C60C30'))

    # Update layout to add axis lines, light grey grids for alternate points, make background transparent
    # and disable zooming, panning, and scrolling
    fig.update_layout(
        width=1100,
        height=400,
        title={'x': 0.5},
        paper_bgcolor="rgba(255, 255, 255, 0.0)",
        plot_bgcolor="rgba(255, 255, 255, 0.0)",
        font={'color': "black", 'family': "Arial"},
        xaxis=dict(
            showline=True, linecolor='#707070', showgrid=False, gridcolor='black', dtick='D2',
            fixedrange=True  # Disable zooming and panning on x-axis
        ),
        yaxis=dict(
            showline=True, linecolor='#707070', showgrid=True, gridcolor='rgba(0, 0, 0, 0.1)', dtick=2,
            fixedrange=True,  # Disable zooming and panning on y-axis
            range = [0, max(filtered_hours) + 2]  # Set the range for y-axis to start from 0

        ),
        dragmode=False  # Disable drag interactions
    )

    # Disable scroll and zoom interactions
    fig.update_layout(xaxis=dict(fixedrange=True), yaxis=dict(fixedrange=True))
    fig.update_layout(dragmode=False)

    return fig

def create_donut_chart(employee_dict_dashboard):
    employee_data = employee_dict_dashboard
    status = employee_data['Status']

    # Extract attendance status counts
    status_counts = {
        'Present': status.count('P'),
        'Present Half Day': status.count('P1/2'),
        'Holiday': status.count('HO'),
        'Absent': status.count('A'),
        'WOP': status.count('WOP'),
        'WOS': status.count('WOS')
    }

    # Create labels with values in brackets
    labels = [f"{key} ({value})" for key, value in status_counts.items()]
    values = list(status_counts.values())
    colors = ['#143D60', '#3282B8', '#FCC737', '#8E1616', '#0CAFFF', '#1B262C']  # Custom colors for each status

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=.4,
        marker=dict(colors=colors)  # Add white outlines
    )])
    fig.update_traces(textinfo='percent+label', textposition='inside', insidetextorientation='radial')

    fig.update_layout(
        width=500,  # Set the width of the chart
        height=470,  # Set the height of the chart
        title_text=f'Attendance Status',
        title_x=0.5,
        title_font=dict(size=20, color='black'),
        annotations=[dict(text='', x=0.5, y=0.5, font_size=20, showarrow=False)],
        margin=dict(t=70, b=10, l=40, r=50),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=12)
        ),
        paper_bgcolor='rgba(255, 255, 255, 0.0)',# Remove background color
        plot_bgcolor='rgba(255, 255, 255, 0.0)'  # Remove plot background color
    )

    return fig

def create_combined_barchart(employee_dict_dashboard):


    employee_data = employee_dict_dashboard

    days = employee_data['Days']
    Half_Day_Mapping = employee_data['halfDayMap']
    earlyLeaveMap = employee_data['earlyLeaveMap']
    lateMark = employee_data['lateMark']
    absenteeMap = employee_data['absenteeMap']

    # Format the days to show only the day of the month
    formatted_days = [day.split()[0].zfill(2) for day in days]

    # Create subplots with no vertical spacing
    fig = make_subplots(
        rows=4, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.05,
    )

    # Add Half Day Mapping heatmap
    fig.add_trace(go.Heatmap(
        z=[Half_Day_Mapping],
        x=formatted_days,
        y=['Half Day Mapping'],
        colorscale=[[0, '#D0DDD0'], [1, '#2E5077']],  # Light grey for 0 and blue for 1
        showscale=False,
        xgap=1.5,  # Gap between boxes along x-axis
        ygap=0,  # No gap between boxes along y-axis
        text=[f"Day: {day}, Half Day Mapping: {mark}" for day, mark in zip(formatted_days, Half_Day_Mapping)],
        hoverinfo="text",
        zmin=0, zmax=1  # Ensure the colorscale does not interpolate
    ), row=1, col=1)

    # Add Late Mark heatmap
    fig.add_trace(go.Heatmap(
        z=[lateMark],
        x=formatted_days,
        y=['Late Mark'],
        colorscale=[[0, '#D0DDD0'], [1, '#1B262C']],  # Light grey for 0 and blue for 1
        showscale=False,
        xgap=1.5,  # Gap between boxes along x-axis
        ygap=0.5,  # No gap between boxes along y-axis
        text=[f"Day: {day}, Late Mark: {mark}" for day, mark in zip(formatted_days, lateMark)],
        hoverinfo="text",
        zmin=0, zmax=1  # Ensure the colorscale does not interpolate
    ), row=2, col=1)

    # Add Early Leave Map heatmap
    fig.add_trace(go.Heatmap(
        z=[earlyLeaveMap],
        x=formatted_days,
        y=['Early Leave Map'],
        colorscale=[[0, '#D0DDD0'], [1, '#151515']],  # Light grey for 0 and blue for 1
        showscale=False,
        xgap=1.5,  # Gap between boxes along x-axis
        ygap=0.5,  # No gap between boxes along y-axis
        text=[f"Day: {day}, Early Leave: {mark}" for day, mark in zip(formatted_days, earlyLeaveMap)],
        hoverinfo="text",
        zmin=0, zmax=1  # Ensure the colorscale does not interpolate
    ), row=3, col=1)

    # Add Absentee Map heatmap
    fig.add_trace(go.Heatmap(
        z=[absenteeMap],
        x=formatted_days,
        y=['Absentee Map'],
        colorscale=[[0, '#D0DDD0'], [1, '#2E5077']],  # Light grey for 0 and tomato red for 1
        showscale=False,
        xgap=1.5,  # Gap between boxes along x-axis
        ygap=0.5,  # No gap between boxes along y-axis
        text=[f"Day: {day}, Absentee: {mark}" for day, mark in zip(formatted_days, absenteeMap)],
        hoverinfo="text",
        zmin=0, zmax=1  # Ensure the colorscale does not interpolate
    ), row=4, col=1)

    # Update layout for the combined chart
    fig.update_layout(
        xaxis=dict(
            tickmode='array',
            tickvals=formatted_days,
            ticktext=formatted_days,
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide x-axis line
            showgrid=False,
            fixedrange=True  # Make x-axis static
        ),
        xaxis2=dict(
            tickmode='array',
            tickvals=formatted_days,
            ticktext=formatted_days,
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide x-axis line
            showgrid=False,
            fixedrange=True  # Make x-axis static
        ),
        xaxis3=dict(
            tickmode='array',
            tickvals=formatted_days,
            ticktext=formatted_days,
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide x-axis line
            showgrid=False,
            fixedrange=True  # Make x-axis static
        ),
        xaxis4=dict(
            tickmode='array',
            tickvals=formatted_days,
            ticktext=formatted_days,
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide x-axis line
            showgrid=False,
            fixedrange=True  # Make x-axis static
        ),
        yaxis=dict(
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide y-axis line
            showgrid=False,
            fixedrange=True,  # Make y-axis static
            tickvals=[]  # Hide y-axis labels
        ),
        yaxis2=dict(
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide y-axis line
            showgrid=False,
            fixedrange=True,  # Make y-axis static
            tickvals=[]  # Hide y-axis labels
        ),
        yaxis3=dict(
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide y-axis line
            showgrid=False,
            fixedrange=True,  # Make y-axis static
            tickvals=[]  # Hide y-axis labels
        ),
        yaxis4=dict(
            gridcolor='black',  # Dark black coloured grids
            zeroline=False,  # Hide y-axis line
            showgrid=False,
            fixedrange=True,  # Make y-axis static
            tickvals=[]  # Hide y-axis labels
        ),
        font={'color': "black", 'family': "Arial"},
        paper_bgcolor="rgba(225, 225, 225, 0)",  # Transparent background
        plot_bgcolor="rgba(225, 225, 225, 0)",  # Transparent plot background
        height=230,  # Set the height of the plot to accommodate the fourth heatmap
        width=1750,  # Set a width for the plot to maintain aspect ratio
        margin=dict(l=180, r=90, b=10, t=10),  # Set margins (left, right, bottom, top)
        annotations=[
            dict(
                x= -0.045,
                y=0.95,
                xref='paper',
                yref='paper',
                text='Half Days',
                showarrow=False,
                font=dict(size=13),
                textangle=0
            ),
            dict(
                x=-0.052,
                y=0.68,
                xref='paper',
                yref='paper',
                text='Late Marks',
                showarrow=False,
                font=dict(size=13),
                textangle=0
            ),
            dict(
                x=-0.060,
                y=0.39,
                xref='paper',
                yref='paper',
                text='Early Leaves',
                showarrow=False,
                font=dict(size=13),
                textangle=0
            ),
            dict(
                x=-0.055,
                y=0.08,
                xref='paper',
                yref='paper',
                text='Absentees',
                showarrow=False,
                font=dict(size=13),
                textangle=0
            )
        ]
    )

    return fig

def create_overtime_barchart(employee_dict_dashboard):
    def format_time_in_hours_and_minutes(hours_decimal):
        """Convert decimal hours to HH:MM format"""
        hours = int(hours_decimal)
        minutes = int((hours_decimal - hours) * 60)
        return f"{hours:02d}:{minutes:02d}"

    def format_time_in_hours_and_minutes_text(hours_decimal):
        """Convert decimal hours to HHhr:MMmins format"""
        hours = int(hours_decimal)
        minutes = int((hours_decimal - hours) * 60)
        return f"{hours:02d}hr:{minutes:02d}mins"


    employee_data = employee_dict_dashboard
    days = employee_data['Days']
    overtimes = employee_data['overTime']

    # Initialize lists for all dates and corresponding overtimes
    all_days = []
    all_overtimes = []
    bar_colors = []

    # Populate all_days with formatted day numbers and all_overtimes with corresponding overtime values
    for day, overtime in zip(days, overtimes):
        day_number = datetime.strptime(day, '%d %B %Y, %A').day
        all_days.append(f'{day_number:02d}')

        hours, minutes = map(int, overtime.split(':'))
        total_minutes = hours * 60 + minutes
        overtime_hours = total_minutes / 60  # Convert to hours
        all_overtimes.append(overtime_hours)

        # Determine bar color based on overtime
        if overtime_hours >= 1:
            bar_colors.append("#143D60")  # Green color for more than 1 hour
        else:
            bar_colors.append("#8E1616")  # Dark grey color for less than 1 hour

    text_outside_bars = [format_time_in_hours_and_minutes(ot) if ot > 0 else "" for ot in all_overtimes]
    hover_texts = [format_time_in_hours_and_minutes_text(ot) if ot > 0 else "" for ot in all_overtimes]

    # Create the bar chart
    fig = go.Figure(data=[
        go.Bar(
            x=all_days,
            y=all_overtimes,
            marker_color=bar_colors,
            text=text_outside_bars,
            hovertext=hover_texts,
            hoverinfo='text',
            textposition='outside'
        )
    ])

    fig.update_layout(
        title=f"Daily Overtime Overview",
        title_x=0.5,  # Center the title
        xaxis_title="Date",
        yaxis_title="Overtime (Hours)",
        xaxis=dict(
            showline=True,
            showgrid=False,
            showticklabels=True,
            linecolor='black',
            linewidth=2,
            ticks='outside',
            tickfont=dict(
                family='Arial',
                size=12,
                color='black',
            ),
            tickangle=0  # No rotation for tick labels
        ),
        yaxis=dict(
            showline=True,
            showgrid=False,
            showticklabels=True,
            linecolor='black',
            linewidth=2,
            ticks='outside',
            tickfont=dict(
                family='Arial',
                size=12,
                color='black',
            ),
        ),
        paper_bgcolor="rgba(225, 225, 225, 0.0)",
        plot_bgcolor="rgba(225, 225, 225, 0.0)",
        font=dict(
            color="black",
            family="Arial"
        ),
        margin=dict(l=125, r=30, t=70, b=70),  # Adjust bottom margin for tick labels
        width=1100,
        height=450,
    )

    # Disable default plotly interactions
    fig.update_layout(
        dragmode=False,
        newshape=dict(line_color='cyan'),
        modebar=dict(
            remove=['zoom', 'pan', 'select', 'lasso2d', 'zoomIn2d', 'zoomOut2d', 'autoScale2d', 'resetScale2d',
                    'hoverClosestCartesian', 'hoverCompareCartesian']
        )
    )

    return fig


def generate_star_rating_html(employee_dict):
    """
    Generates an HTML-friendly star rating graphic based on the value of the 'stars' key in the input dictionary.

    :param star_dict: A dictionary containing a key 'stars' with an integer value representing the number of stars.
    :return: A string containing HTML code for the star rating graphic.
    """
    adherenceRatioValue = employee_dict['reportMetric'].get("adherenceRatio", 0)
    workDeficitRatioValue = employee_dict['reportMetric'].get("workDeficitRatio", 0)
    adjustAbsenteeValue = employee_dict['reportMetric'].get("adjustedAbsenteeRate", 0)

    adherenceRatioStar = employee_dict['reportMetric'].get("adherenceRatioStar", 0)
    workDeficitRatioStar = employee_dict['reportMetric'].get("workDeficitRatioStar", 0)
    adjustAbsenteeStar = employee_dict['reportMetric'].get("adjustedAbsenteeRateStar", 0)

    # Generate star rating graphic
    full_star = "★"
    empty_star = "☆"
    max_stars = 5

    adherence_star_rating = full_star * adherenceRatioStar + empty_star * (max_stars - adherenceRatioStar)
    work_deficit_star_rating = full_star * workDeficitRatioStar + empty_star * (max_stars - workDeficitRatioStar)
    adjusted_absantee_star_rating = full_star * adjustAbsenteeStar + empty_star * (max_stars - adjustAbsenteeStar)

    # Generate HTML-friendly output
    html_output = f"""

    <div class="ratings-container">
        <div class="punctuality-rating">
            <h4>Punctuality Rating</h4>
            <div class="star-rating" style="font-size: 45px; color: #FCC737;"> <!-- For stars -->
                {adherence_star_rating}
            </div>
            <div class="number-rating" style="font-size: 25px; color: #1B262C;"> <!-- For numbers -->
                {adherenceRatioValue}
            </div>
        </div>

        <div class="overtime-to-incomplete-rating">
            <h4>Overtime to Incomplete Rating</h4>
            <div class="star-rating" style="font-size: 45px; color: #FCC737;"> <!-- For stars -->
                {work_deficit_star_rating}
            </div>
            <div class="number-rating" style="font-size: 25px; color: #1B262C;"> <!-- For numbers -->
                {workDeficitRatioValue}
            </div>
        </div>

        <div class="absentism-rating">
            <h4>Absentism Rate</h4>
            <div class="star-rating" style="font-size: 45px; color: #FCC737;"> <!-- For stars -->
                {adjusted_absantee_star_rating}
            </div>
            <div class="number-rating" style="font-size: 25px; color: #1B262C;"> <!-- For numbers -->
                {adjustAbsenteeValue}
            </div>
        </div>
    </div>
    """

    return html_output
