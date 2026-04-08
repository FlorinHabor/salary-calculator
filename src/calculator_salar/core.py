import pandas as pd
import openpyxl

global headers

def parse_excel_file(file_path):
    global headers
    print(f"Parsing file: {file_path}")
    print(file_path)
    
    df = pd.read_excel(file_path, header = None)
    headers = find_headers_row(df)
    return calculate_hours(df, headers)

def parse_time_cell(cell):
    """Convert hh:mm string or float to hours as float."""
    if isinstance(cell, str):
        try:
            cell = cell.replace(" ", "")
            h, m = map(int, cell.split(":"))
            return h + m / 60
        except:
            return 0
    elif pd.notna(cell):
        return float(cell)
    else:
        return 0

def calculate_hours(df, headers):
    """
    Returns:
        dict: {employee_name: {"worked_days": int, "normal_hours": float, "overtime": float}}
    """
    employee_data = {}

    first_name_col = headers["First Name"]
    total_work_col = headers["Total Work Hours"]
    overtime_col = headers["Total Overtime"]

    for row_index in range(headers["start_row"] + 1, len(df)):
        row = df.iloc[row_index]

        name = row[first_name_col]
        if pd.isna(name) or str(name).strip() == "":
            continue

        if name not in employee_data:
            employee_data[name] = {
                "worked_days": 0,
                "normal_hours": 0.0,
                "overtime": 0.0
            }

        total_work = parse_time_cell(row[total_work_col])
        overtime = parse_time_cell(row[overtime_col])
        normal_hours = max(total_work - overtime, 0)

        if total_work >= 1:
            employee_data[name]["worked_days"] += 1

        employee_data[name]["normal_hours"] += normal_hours
        employee_data[name]["overtime"] += overtime

    return employee_data

def find_headers_row(df):
    global headers
    rows_to_search  = 15
    column_to_search = 0
    headers = {}
    for row_index in range(rows_to_search):
        cell = df.iloc[row_index, column_to_search]
        if isinstance(cell, str) and cell.strip().lower() == "first name":
            headers['start_row'] = row_index
            print(f"Found first name in row {row_index}")
            print(f"Header row: {df.shape[1]}")
            for col_index in range( df.shape[1]):
                header_cell = df.iloc[row_index, col_index]
                headers[header_cell] = col_index
            return headers
    return None

def get_data_for_employee(employee_name, file_path):
    global headers
    df = pd.read_excel(file_path, header = None)
    employee = []
    for row_index in range(len(df)):
        row = df.iloc[row_index,0]
        if str(row).strip() == employee_name:
            if(parse_time_cell(df.iloc[row_index,headers['Total Work Hours']]) != 0):
                employee.append({
                    "date": df.iloc[row_index,headers['Date']],
                    "required_hours": 8,
                    "total_hours":  parse_time_cell(df.iloc[row_index,headers['Total Work Hours']]),
                    "normal_hours": max(parse_time_cell(df.iloc[row_index,headers['Total Work Hours']]) - parse_time_cell(df.iloc[row_index,headers['Total Overtime']]), 0),
                    "missing_hours": max(0, 8 - parse_time_cell(df.iloc[row_index,headers['Total Work Hours']]))                ,
                    "overtime": parse_time_cell(df.iloc[row_index,headers['Total Overtime']])
                })
    return employee