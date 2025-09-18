# -*- coding: utf-8 -*-
"""
Created on Wed Jul  5 17:22:56 2023
Updated on Sat Sep  6 2025

@author: shank
"""

import pandas as pd
import time
import requests
import tkinter as tk
from tkcalendar import Calendar
import pytz


def find_detail(custom_fields, detail):
    """Finds a custom field value by its name."""
    try:
        for field in custom_fields:
            if field["name"].strip().lower() == detail.strip().lower():
                return field.get("value")
    except Exception:
        return None
    return None


def get_option_value(custom_fields, detail):
    """Gets the option value from the list of custom fields."""
    try:
        for field in custom_fields:
            if field["name"].strip().lower() == detail.strip().lower():
                if field.get("value") is None:
                    return None
                return field["type_config"]["options"][field["value"]]["name"]
    except Exception:
        return None
    return None


def get_data_for_these_dates():
    start_date = start_cal.selection_get()
    end_date = end_cal.selection_get()
    start_date = int(time.mktime(start_date.timetuple()))
    end_date = int(time.mktime(end_date.timetuple())) + 86399  # end of day

    ### Extracting all status names from the space
    space_id = "3565019"
    url = f"https://api.clickup.com/api/v2/space/{space_id}"

    headers = {
        "Content-Type": "application/json",
        "Authorization": "pk_3326657_EOM3G6Z3CKH2W61H8NOL5T7AGO9D7LNN"
    }

    response = requests.get(url, headers=headers)
    data = response.json()
    statuses_list = data.get("statuses", [])
    status_values = [status["status"] for status in statuses_list]

    list_id = "11943493"  # Customer Ticketing System
    url = f"https://api.clickup.com/api/v2/list/{list_id}/task"

    query = {
        "archived": "false",
        "page": "0",
        "date_created_gt": str(start_date) + '000',
        "date_created_lt": str(end_date) + '000',
        "statuses": status_values
    }

    all_tasks = []
    while True:
        response = requests.get(url, headers=headers, params=query)
        data = response.json()

        all_tasks.extend(data.get("tasks", []))

        if data.get("last_page", True):
            break
        query['page'] = str(int(query['page']) + 1)

    # Construct dataframe
    df = pd.DataFrame({
        "Issue Description": [task.get("name", "") for task in all_tasks],
        "status": [task.get("status", {}).get("status", "") for task in all_tasks],
        "priority": [task["priority"]["priority"] if task.get("priority") else ""
                     for task in all_tasks],
        "timeSpent": [task.get("time_spent", 0) for task in all_tasks],
        "username": [task["assignees"][0]["username"] if task.get("assignees") else ""
                     for task in all_tasks],
        "createdDate": [task.get("date_created") for task in all_tasks],
        "custName": [find_detail(task.get("custom_fields", []), "Cust. Name") for task in all_tasks],
        "custPhone": [find_detail(task.get("custom_fields", []), "Cust. Phone") for task in all_tasks],
        "custEmail": [find_detail(task.get("custom_fields", []), "Cust Email") for task in all_tasks],
        "channel": [get_option_value(task.get("custom_fields", []), "Channel") for task in all_tasks],
        "country": [get_option_value(task.get("custom_fields", []), "country") for task in all_tasks],
        "website": [get_option_value(task.get("custom_fields", []), "Website") for task in all_tasks],
        "product": [get_option_value(task.get("custom_fields", []), "Product") for task in all_tasks],
        "course": [get_option_value(task.get("custom_fields", []), "Course") for task in all_tasks],
        "Type of Issue": [get_option_value(task.get("custom_fields", []), "Type of Issue") for task in all_tasks],
        "Type of Query": [get_option_value(task.get("custom_fields", []), "Type of Query") for task in all_tasks],
        "Age Group": [get_option_value(task.get("custom_fields", []), "Age Group") for task in all_tasks],
        "reportedDate": [find_detail(task.get("custom_fields", []), "Reported Date") for task in all_tasks],
        "resolvedDate": [find_detail(task.get("custom_fields", []), "Resolved Date") for task in all_tasks],
        "responseGiven": [find_detail(task.get("custom_fields", []), "Response Given") for task in all_tasks],
        "Task ID": [task.get("id") for task in all_tasks]
    })

    def change_date_format(input_date):
        ist = pytz.timezone('Asia/Kolkata')
        formatted_date = input_date.dt.tz_localize(pytz.utc).dt.tz_convert(ist)
        output_format = '%m/%d/%Y, %I:%M:%S %p %Z'
        return formatted_date.dt.strftime(output_format)

    # Convert to datetime safely
    for col in ["createdDate", "reportedDate", "resolvedDate"]:
        df[col] = pd.to_datetime(df[col], unit="ms", errors="coerce")

    df["TAT"] = (df["resolvedDate"] - df["reportedDate"]).dt.days

    df["createdDate"] = change_date_format(df["createdDate"].dropna())
    df["reportedDate"] = change_date_format(df["reportedDate"].dropna())
    df["resolvedDate"] = change_date_format(df["resolvedDate"].dropna())

    # Time spent to h m format
    df["timeSpent"] = df["timeSpent"].apply(
        lambda x: f"{int((x/1000) / (60 * 60))}h {int((x/1000) % (60 * 60)) // 60}m" if x else "0h 0m"
    )

    df['Task URL'] = 'https://app.clickup.com/t/' + df['Task ID'].astype(str)

    # Reorder some cols
    task_id_col = df.pop('Task ID')
    df.insert(len(df.columns) - 1, 'Task ID', task_id_col)

    reportedDate_col = df.pop('reportedDate')
    df.insert(len(df.columns) - 1, 'reportedDate', reportedDate_col)

    filename = "CustomerTickets.xlsx"
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    worksheet = writer.sheets['Sheet1']

    # Write clickable Task ID
    for row_num, value in enumerate(df['Task ID'], start=1):
        if pd.isna(value):
            continue
        url = f'https://app.clickup.com/t/{value}'
        worksheet.write_url(row_num, df.columns.get_loc('Task ID'), url, string=str(value))

    writer.close()

    df.to_csv("CustomerTickets.csv", index=False)


# GUI
root = tk.Tk()
root.title("Customer Ticket Resolution - Analytics.")

window_width = 500
window_height = 550
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

root.configure(background="lightblue")

# Start Date
start_frame = tk.Frame(root)
start_frame.pack(pady=10)
start_label = tk.Label(start_frame, text="Start\nDate:", bg="black", fg="white")
start_label.pack(side="left")
start_cal = Calendar(start_frame, selectmode="day", date_pattern="yyyy-mm-dd")
start_cal.pack(side="left")

# End Date
end_frame = tk.Frame(root)
end_frame.pack(pady=10)
end_label = tk.Label(end_frame, text="End\nDate:", bg="black", fg="white")
end_label.pack(side="left")
end_cal = Calendar(end_frame, selectmode="day", date_pattern="yyyy-mm-dd")
end_cal.pack(side="left")

submit_button = tk.Button(root, text="Submit", command=get_data_for_these_dates)
submit_button.pack(pady=10)

output_label = tk.Label(root,
                        text="Note: Excel & CSV output saved in this folder.",
                        font=("Times", 12, "bold"),
                        bg="red", fg="yellow")
output_label.pack()

footer_label = tk.Label(root, text="Version 1.5 (6th Sept 2025)", relief=tk.RAISED, anchor=tk.W)
footer_label.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()
