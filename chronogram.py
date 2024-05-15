import pandas as pd
import re
from datetime import timedelta, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell

def allocateTasksToWeeks(milestones_tasks):
    chronogram = []
    last_end_week = 0  # Track the end week of the last task

    for milestone_name, tasks in milestones_tasks:
        if len(chronogram) > 0:
            last_end_week += 1  # Ensure the new milestone starts on a new week

        colWeekHours = [40] * (last_end_week + 1)  # Ensure we have enough weeks to start
        milestone_rows = []

        for task in tasks:
            initial_task_hours = task  # Store the initial task hours for printing
            weeks = len(colWeekHours)
            taskRow = ['_'] * weeks  # Current row times the weeks needed
            while task > 0:  # While the task has hours left to assign
                for i in range(last_end_week, len(colWeekHours)):
                    if task <= colWeekHours[i]:  # Task needs fewer hours than available in current work week
                        colWeekHours[i] -= task
                        taskRow[i] = 'X'
                        task = 0  # Task hours fully allocated
                        last_end_week = max(last_end_week, i)  # Update the end week
                        print(f"Milestone: {milestone_name}, Task Hours: {initial_task_hours}, Remaining Task: {task}, Week: {i+1}, Assigned: X")
                        break
                    else:  # Task needs more hours than available in current work week
                        if colWeekHours[i] > 0:
                            task -= colWeekHours[i]
                            taskRow[i] = 'X'
                            colWeekHours[i] = 0
                            last_end_week = max(last_end_week, i)  # Update the end week
                            print(f"Milestone: {milestone_name}, Task Hours: {initial_task_hours}, Remaining Task: {task}, Week: {i+1}, Assigned: X")

                # If task still has hours left not allocated, add new week
                if task > 0:
                    colWeekHours.append(40)
                    taskRow.append('_')  # Extend task row for the new week

            milestone_rows.append(taskRow)

        # Add milestone rows to the chronogram
        chronogram.extend(milestone_rows)
        # Add an empty row after each milestone except the last one
        if milestone_name != milestones_tasks[-1][0]:
            chronogram.append([''] * len(colWeekHours))

    return chronogram

# Global storage for all week dates
all_week_dates = []

def validate_date(date_text):
    try:
        datetime.strptime(date_text, '%m/%d')
        return True
    except ValueError:
        return False

def add_task_dates(chronogram, start_date, ws, year, num_weeks, task_row_mapping, task_milestone_mapping, milestone_row_mapping, row_offset=4):
    if not start_date:
        return None

    current_milestone = None
    global_start_date = datetime.strptime(f"{start_date}/{year}", "%m/%d/%Y")  # Initial global start date for all tasks

    def get_next_available_date(date, used_dates):
        while date in used_dates:
            date += timedelta(days=7)
        return date

    used_start_dates = []
    used_end_dates = []

    print("From add_task_dates", end='\n')

    week_dates = get_week_dates(start_date, num_weeks, year)

    milestone_start_dates = {}
    milestone_end_dates = {}

    for index, task_row in enumerate(chronogram):
        # Check if the row is empty
        if set(task_row) == {''}:
            print(f"Skipping empty row {index + 1}")
            continue

        milestone_name = task_milestone_mapping[index]
        if current_milestone != milestone_name:
            current_milestone = milestone_name
            print(f"\nProcessing tasks for Milestone: {current_milestone}")
            if current_milestone and used_end_dates:  # Adjust start date only if we move to a new milestone and have used_end_dates
                global_start_date = max(used_end_dates) + timedelta(days=1)  # Start the next milestone after the last task end date
                print(f"Adjusted global_start_date for new milestone: {global_start_date.strftime('%m/%d/%Y')}")

        x_indices = [i for i, x in enumerate(task_row) if x == 'X']
        if x_indices and len(week_dates) > x_indices[0]:
            start_week_range, _ = week_dates[x_indices[0]]
            end_week_range, _ = week_dates[x_indices[-1]]
            start_week_range = start_week_range.split(' - ')[0]
            end_week_range = end_week_range.split(' - ')[1]

            task_start_date = datetime.strptime(f"{start_week_range}/{year}", "%d/%b/%Y")
            task_end_date = datetime.strptime(f"{end_week_range}/{year}", "%d/%b/%Y")

            task_start_date = get_next_available_date(task_start_date, used_start_dates)
            duration_days = (task_end_date - task_start_date).days
            task_end_date = task_start_date + timedelta(days=duration_days)
            task_end_date = get_next_available_date(task_end_date, used_end_dates)

            used_start_dates.append(task_start_date)
            used_end_dates.append(task_end_date)

            # Write the start and end dates only if the cell is not part of a merged cell range
            if not isinstance(ws.cell(row=task_row_mapping[index], column=4), MergedCell):
                ws.cell(row=task_row_mapping[index], column=4, value=task_start_date.strftime("%m/%d"))
            if not isinstance(ws.cell(row=task_row_mapping[index], column=5), MergedCell):
                ws.cell(row=task_row_mapping[index], column=5, value=task_end_date.strftime("%m/%d"))

            # Track start and end dates for milestones
            if milestone_name not in milestone_start_dates or task_start_date < milestone_start_dates[milestone_name]:
                milestone_start_dates[milestone_name] = task_start_date
            if milestone_name not in milestone_end_dates or task_end_date > milestone_end_dates[milestone_name]:
                milestone_end_dates[milestone_name] = task_end_date

            # Fill in the appropriate columns with orange cells
            for i, (date_range, _) in enumerate(week_dates, start=6):
                start_week_str, end_week_str = date_range.split(' - ')
                start_week = datetime.strptime(f"{start_week_str}/{year}", "%d/%b/%Y")
                end_week = datetime.strptime(f"{end_week_str}/{year}", "%d/%b/%Y")

                if task_start_date <= end_week and task_end_date >= start_week:
                    task_cell = ws.cell(row=task_row_mapping[index], column=i)
                    task_cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
                    print(f"Filling cell: Row {task_row_mapping[index]}, Column {i} for date range {start_week.strftime('%m/%d')} to {end_week.strftime('%m/%d')}")
            print(f"Task [{index}]: Start - {task_start_date.strftime('%m/%d')} End - {task_end_date.strftime('%m/%d')}")
        else:
            print(f"Skipping row {index + 1} as it contains no tasks or insufficient weeks.")

    # Add milestone start and end dates to the milestone rows and fill with orange color
    for milestone_name, start_date in milestone_start_dates.items():
        end_date = milestone_end_dates[milestone_name]
        milestone_row = milestone_row_mapping[milestone_name]
        ws.cell(row=milestone_row, column=4, value=start_date.strftime("%m/%d"))
        ws.cell(row=milestone_row, column=5, value=end_date.strftime("%m/%d"))

        # Fill milestone row cells with orange color for the weeks it spans
        for i, (date_range, _) in enumerate(week_dates, start=6):
            start_week_str, end_week_str = date_range.split(' - ')
            start_week = datetime.strptime(f"{start_week_str}/{year}", "%d/%b/%Y")
            end_week = datetime.strptime(f"{end_week_str}/{year}", "%d/%b/%Y")

            if start_date <= end_week and end_date >= start_week:
                milestone_cell = ws.cell(row=milestone_row, column=i)
                milestone_cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

    return None

def calculate_total_weeks(chronogram):
    """
    Calculate the total number of weeks required based on the chronogram data.
    """
    max_length = max(len(row) for row in chronogram if set(row) != {''})
    return max_length

def adjust_column_settings(ws, start_col_index, num_weeks):
    column_widths = {
        'B': 5,  # Tasks column
        'C': 30,  # Activity column
        'D': 12,  # Start Date column
        'E': 12,  # End Date column
    }

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Set the width of the week date headers dynamically
    week_date_col_width = 16  # Adequate width to display full date ranges
    for i in range(num_weeks):
        col_letter = get_column_letter(start_col_index + i)
        ws.column_dimensions[col_letter].width = week_date_col_width

    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=2, max_col=3):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

def process_final_week_ranges():
    global all_week_ranges
    print("")
    return all_week_ranges

# Global variables to manage state
current_milestone = None
last_milestone_end_date = None
milestone_start_date = None
all_week_ranges = []  # This will collect all unique week ranges
milestone_count = 0
current_milestone_count = 1
last_activity = None

def get_week_dates(start_date, num_weeks, year, milestone_name=None, last_end_dates=None, is_last_task=False):
    global last_milestone_end_date, current_milestone, milestone_start_date, all_week_ranges, current_milestone_count, milestone_count

    week_dates = []
    start_dates = []

    if last_milestone_end_date is not None and milestone_name != current_milestone:
        new_start_date = last_milestone_end_date + timedelta(days=1)
        start_dates = [new_start_date]
        milestone_start_date = new_start_date
    elif milestone_name == current_milestone and milestone_start_date:
        start_dates = [milestone_start_date]
    elif last_end_dates is not None:
        start_dates = [last_end_date + timedelta(days=1) for last_end_date in last_end_dates]
    if not start_dates:
        start_dates = [datetime.strptime(f"01/01/{year}", "%m/%d/%Y")]
        milestone_start_date = start_dates[0]

    print(f"Calculating week dates from start date: {start_dates[0].strftime('%m/%d/%Y')} for {num_weeks} weeks")
    
    current_dates = start_dates

    for i in range(num_weeks):
        end_dates = [current_date + timedelta(days=6) for current_date in current_dates]
        current_week_ranges = [f"{current_date.strftime('%d/%b')} - {end_date.strftime('%d/%b')}" for current_date, end_date in zip(current_dates, end_dates)]
        week_dates.extend([(week_range, current_date.year) for week_range, current_date in zip(current_week_ranges, current_dates)])
        all_week_ranges.extend([(week_range, current_date.year) for week_range, current_date in zip(current_week_ranges, current_dates)])
        current_dates = [end_date + timedelta(days=1) for end_date in end_dates]

    if milestone_name:
        last_milestone_end_date = end_dates[-1] if end_dates else None
        current_milestone = milestone_name

    if current_milestone_count == milestone_count and milestone_name == current_milestone:
        process_final_week_ranges()

    print("Week dates calculated: ", week_dates)
    return week_dates

def chronogramToExcel(chronogram, year, start_week, activity_names, milestoneNames, filename="chronogram.xlsx"):
    start_col_index = 6
    num_weeks = calculate_total_weeks(chronogram)

    df = pd.DataFrame(chronogram)
    for col in range(start_col_index - 1):
        df.insert(col, 'Empty{}'.format(col), [''] * df.shape[0])
    df.to_excel(filename, index=False, header=False)

    if not year:
        week_years = set([date_info[1] for date_info in week_dates])
        if len(week_years) == 1:
            year = week_years.pop()
        else:
            year = min(week_years)

    wb = Workbook()
    ws = wb.active

    ###########################Headers area

    row_offset = 5

    milestone_index = 0
    task_index = 1
    activity_index = 0
    task_row_mapping = {}
    task_milestone_mapping = {}
    milestone_row_mapping = {}

    new_chronogram = []
    new_activity_names = []

    # Insert milestone rows and adjust mappings
    for index, row in enumerate(chronogram):
        if set(row) == {''}:
            milestone_index += 1
            task_index = 1
            continue

        if milestoneNames[milestone_index] not in milestone_row_mapping:
            milestone_row_mapping[milestoneNames[milestone_index]] = len(new_chronogram) + row_offset
            new_chronogram.append([''] * len(row))
            new_activity_names.append(milestoneNames[milestone_index])

        task_label = f"Task {milestone_index + 1}.{task_index}"
        task_row_mapping[len(new_chronogram)] = len(new_chronogram) + row_offset
        task_milestone_mapping[len(new_chronogram)] = milestoneNames[milestone_index]
        task_index += 1

        new_chronogram.append(row)
        if activity_index < len(activity_names):
            new_activity_names.append(activity_names[activity_index])
            activity_index += 1

    # Write milestones and tasks to the worksheet
    milestone_counter = 0
    for index, row in enumerate(new_chronogram):
        excel_row = row_offset + index

        if set(row) == {''}:
            ws.cell(row=excel_row, column=2, value=f"Task {milestone_counter + 1}")
            ws.merge_cells(start_row=excel_row, start_column=2, end_row=excel_row, end_column=2)
            task_cell = ws.cell(row=excel_row, column=2)
            task_cell.alignment = Alignment(horizontal='center', vertical='center')
            task_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
            task_cell.font = Font(color="FFFFFF", bold=True)

            ws.cell(row=excel_row, column=3, value=milestoneNames[milestone_counter])
            ws.merge_cells(start_row=excel_row, start_column=3, end_row=excel_row, end_column=3)
            cell = ws.cell(row=excel_row, column=3)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)

            milestone_counter += 1
        else:
            if index in task_row_mapping:
                task_excel_row = task_row_mapping[index]
                ws.cell(row=task_excel_row, column=2, value=f"Task {milestone_counter}.{task_index}")
                ws.cell(row=task_excel_row, column=3, value=new_activity_names[index])

                for col_index, value in enumerate(row, start=start_col_index):
                    task_cell = ws.cell(row=task_excel_row, column=col_index)
                    if value == 'X':
                        task_cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

    column_width = 20
    for col in ws.iter_cols(min_col=start_col_index, max_col=ws.max_column, min_row=1, max_row=ws.max_row):
        for cell in col:
            ws.column_dimensions[get_column_letter(cell.column)].width = column_width

    add_task_dates(new_chronogram, start_week, ws, year, num_weeks, task_row_mapping, task_milestone_mapping, milestone_row_mapping)

    week_dates = sorted(set(all_week_ranges), key=lambda x: (x[1], datetime.strptime(x[0].split(' - ')[0], '%d/%b')))

    if not week_dates:
        print("No start date provided. Generating default week dates starting from the first week of the year.")
        week_dates = get_week_dates("01/01", num_weeks, year)

    current_year = week_dates[0][1]
    year_start_col = start_col_index

    for i, (date_range, year_of_week) in enumerate(week_dates, start=start_col_index):
        if current_year != year_of_week:
            ws.merge_cells(start_row=1, start_column=year_start_col, end_row=1, end_column=i - 1)
            primary_cell = ws.cell(row=1, column=year_start_col)
            primary_cell.value = str(current_year)
            primary_cell.alignment = Alignment(horizontal='left', vertical='center')
            primary_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
            primary_cell.font = Font(color="FFFFFF", bold=True)

            current_year = year_of_week
            year_start_col = i

    ws.merge_cells(start_row=1, start_column=year_start_col, end_row=1, end_column=len(week_dates) + start_col_index - 1)
    primary_cell = ws.cell(row=1, column=year_start_col)
    primary_cell.value = str(current_year)
    primary_cell.alignment = Alignment(horizontal='left', vertical='center')
    primary_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    primary_cell.font = Font(color="FFFFFF", bold=True)

    ###########################

    row_offset = 2
    months = {}

    for i, (date_range, year) in enumerate(week_dates, start=start_col_index):
        start_date_str, _ = date_range.split(' - ')
        try:
            start_date = datetime.strptime(start_date_str, "%d/%b")
            month_name = start_date.strftime("%B")
        except ValueError:
            month_name = "Unknown Month"

        if month_name not in months:
            months[month_name] = {'start': i, 'end': i}
        else:
            months[month_name]['end'] = i

        week_cell = ws.cell(row=row_offset + 1, column=i)
        week_cell.value = date_range
        week_cell.alignment = Alignment(horizontal='center')
        week_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
        week_cell.font = Font(color="FFFFFF")

    for month_name, month_range in months.items():
        if month_range['start'] <= month_range['end']:
            ws.merge_cells(start_row=row_offset, start_column=month_range['start'], end_row=row_offset, end_column=month_range['end'])
            month_cell = ws.cell(row=row_offset, column=month_range['start'])
            month_cell.value = month_name
            month_cell.alignment = Alignment(horizontal='center')
            month_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
            month_cell.font = Font(color="FFFFFF")

    ws.merge_cells(start_row=1, start_column=2, end_row=3, end_column=2)
    task_header_cell = ws.cell(row=1, column=2)
    task_header_cell.alignment = Alignment(horizontal='center', vertical='bottom')
    task_header_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    task_header_cell.font = Font(color="FFFFFF", bold=True)
    task_header_cell.value = "Tasks"

    ws.merge_cells(start_row=1, start_column=3, end_row=3, end_column=3)
    activity_header_cell = ws.cell(row=1, column=3)
    activity_header_cell.value = "Activity"
    activity_header_cell.alignment = Alignment(horizontal='center', vertical='bottom')
    activity_header_cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    activity_header_cell.font = Font(color="FFFFFF", bold=True)

    ws.merge_cells(start_row=1, start_column=4, end_row=3, end_column=4)
    ws.merge_cells(start_row=1, start_column=5, end_row=3, end_column=5)

    start_date_header_cell = ws.cell(row=1, column=4)
    end_date_header_cell = ws.cell(row=1, column=5)

    start_date_header_cell.value = "Start Date"
    end_date_header_cell.value = "End Date"

    for cell in [start_date_header_cell, end_date_header_cell]:
        cell.alignment = Alignment(horizontal='center', vertical='bottom')
        cell.fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)

    adjust_column_settings(ws, start_col_index, num_weeks)

    wb.save(filename)
    df.to_csv("chronogram.csv", index=False)

# Ask user for the year for the Gantt Chart
yearInput = input("Add the year for the Gantt Chart (leave empty if using current year): ").strip()
year = int(yearInput) if yearInput else datetime.now().year

# Prompt the user for the starting week, now expecting MM/DD format
start_week = input("Add the starting week (MM/DD) (leave empty if not): ").strip()
while start_week and not validate_date(start_week):
    start_week = input("The format is incorrect. Please use MM/DD format or leave empty: ").strip()

milestoneNames = []
milestonesInput = input("Enter the list of milestones (as comma-separated values), or leave empty: ")
if milestonesInput:
    milestoneNames = [milestone.strip() for milestone in milestonesInput.split(',')]
    # Track the count of milestones
    milestone_count = len(milestoneNames)
else:
    milestone_count = 0  # No milestones entered

activityNames = []
milestones_tasks = []

for index, milestone in enumerate(milestoneNames):
    print(f"Adding tasks for Milestone: {milestone}")

    tasksInput = input(f"Enter the list of tasks for {milestone} (as comma-separated values): ")
    while not tasksInput:
        tasksInput = input(f"Add at least one task for {milestone} (as comma-separated values): ")

    tasks = [task.strip() for task in tasksInput.split(',')]
    while not all(tasks):
        print("Task names can't be empty. Please enter valid task names.")
        tasksInput = input(f"Enter the list of tasks for {milestone} (as comma-separated values): ")
        tasks = [task.strip() for task in tasksInput.split(',')]

    taskHoursInput = input(f"Enter the hours for tasks under {milestone} (as comma-separated values): ")
    while not taskHoursInput:
        taskHoursInput = input(f"Add at least one task hour for {milestone} (as comma-separated values): ")
    hours = [int(x.strip()) for x in re.split(r'[,\s]+', taskHoursInput) if x.strip()]

    milestoneActivityNames = [f"{milestone} - {task}" for task in tasks]
    activityNames.extend(milestoneActivityNames)
    milestones_tasks.append((milestone, hours))

    if index == len(milestoneNames) - 1:
        last_activity = milestoneActivityNames[-1] if milestoneActivityNames else None

# Generate the chronogram from user input
chronogram = allocateTasksToWeeks(milestones_tasks)

# Call the function to save the chronogram to an Excel file
chronogramToExcel(chronogram, year, start_week if start_week.strip() else "", activityNames, milestoneNames, "chronogram.xlsx")
