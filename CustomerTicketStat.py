# -*- coding: utf-8 -*-
"""
Created on Wed Jul  5 17:22:56 2023

@author: shank
"""

import pandas as pd
import time
import requests
import tkinter as tk
from tkcalendar import Calendar

def find_detail(custom_fields, detail):
  """Finds the phone number in the list of custom fields.

  Args:
    custom_fields: A list of dictionaries.

  Returns:
    The phone number.
  """

  cust_detail = None
  try:
      for field in custom_fields:
        if field["name"] == detail:
          cust_detail = field["value"]
          break
  except KeyError as e:
      pass
  return cust_detail

def get_option_value(custom_fields, detail):
  """Gets the option value from the list of custom fields.

  Args:
    custom_fields: A list of dictionaries.

  Returns:
    The option value.
  """

  option_value = None
  try:
      for field in custom_fields:
        if field["name"] == detail:
            option_value = field["type_config"]["options"][field["value"]]['name']
            break
  except KeyError as e:
      pass
  return option_value

def get_data_for_these_dates():
    # start = time.time()
    # QUOTA = 0
    
    start_date = start_cal.selection_get()
    end_date = end_cal.selection_get()
    start_date = int(time.mktime(start_date.timetuple()))
    end_date = int(time.mktime(end_date.timetuple()))+86399 # EOD 11:59 PM instead of 12AM
    
    ### Extracting all status names from the space
    space_id = "3565019"
    url = "https://api.clickup.com/api/v2/space/" + space_id
    
    headers = {
      "Content-Type": "application/json",
      "Authorization": "pk_3326657_EOM3G6Z3CKH2W61H8NOL5T7AGO9D7LNN"
    }
    response = requests.get(url, headers=headers)
    data = response.json()
    # Extract the "statuses" list
    statuses_list = data["statuses"]
    # Extract the "status" values and create a list
    status_values = [status["status"] for status in statuses_list]
    
    list_id = "11943493" #Customer Ticketing System
    url = "https://api.clickup.com/api/v2/list/" + list_id + "/task"
    
    query = {
      "archived": "false",
      "page": "0",
      "date_created_gt": str(start_date)+'000',
      "date_created_lt": str(end_date)+'000',
       "statuses": status_values
    }
    
    # Initialize an empty list to store the concatenated tasks data
    all_tasks = []

    while True:
        response = requests.get(url, headers=headers, params=query)    
        data = response.json()
        
        # Concatenate the tasks data to the list
        all_tasks.extend(data['tasks'])
    
        # Check if it's the last page
        if data['last_page']:
            break
        
        # Increment the page parameter for the next request
        query['page'] = str(int(query['page'])+1)
    
    # Create a Pandas DataFrame.
    df = pd.DataFrame({
        "name": [task["name"] for task in all_tasks],
        # "updatedDate": [task["date_updated"] for task in all_tasks],
        "status": [task["status"]["status"] for task in all_tasks],
        "priority": [task["priority"]["priority"] if task["priority"] is not None else ""
                     for task in all_tasks],
        "timeSpent": [task["time_spent"] if task.get("time_spent") else 0 
                      for task in all_tasks],
        "username": [task["assignees"][0]["username"] if task["assignees"] else ""
                     for task in all_tasks],
        "createdDate": [task["date_created"] for task in all_tasks],
        "custName": [find_detail(task["custom_fields"], "Cust. Name") for task in all_tasks],
        "custPhone": [find_detail(task["custom_fields"], "Cust. Phone") for task in all_tasks],
        "custEmail": [find_detail(task["custom_fields"], "Cust Email") for task in all_tasks],
        "channel": [get_option_value(task["custom_fields"], "Channel") for task in all_tasks],
        "country": [get_option_value(task["custom_fields"], "Country") for task in all_tasks],
        "website": [get_option_value(task["custom_fields"], "Website") for task in all_tasks],
        "product": [get_option_value(task["custom_fields"], "Product") for task in all_tasks],
        "course": [get_option_value(task["custom_fields"], "Course") for task in all_tasks],
        "Type of Issue": [get_option_value(task["custom_fields"], "Type of Issue") for task in all_tasks],
        "resolvedDate": [find_detail(task["custom_fields"], "Resolved Date") for task in all_tasks],
        "responseGiven": [find_detail(task["custom_fields"], "Response Given") for task in all_tasks]
        # "turnaroundHours": [int((int(resolvedDate) - int(createdDate))/3600000) for task in all_tasks]
    })
    
    # Convert the unix time column to date format
    df["createdDate"] = pd.to_datetime(df["createdDate"], unit="ms")
    df["resolvedDate"] = pd.to_datetime(df["resolvedDate"], unit="ms")
    
    # Convert the time column to hours and minutes
    df["timeSpent"] = df["timeSpent"].apply(lambda x: f"{int((x/1000) / (60 * 60))}h {int((x/1000) % (60 * 60)) // 60}m")
    
    # Convert the time columns to datetime format
    df["createdDate"] = pd.to_datetime(df["createdDate"])
    df["resolvedDate"] = pd.to_datetime(df["resolvedDate"])
    # Calculate the difference between the two columns in days and hours
    df["TAT"] = (df["createdDate"] - df["resolvedDate"]).dt.days
    # Save the Pandas DataFrame to an Excel spreadsheet.
    df.to_excel("CustomerTickets.xlsx")

root = tk.Tk()
root.title("Customer Ticket Resolution - Analytics.")

# Increase the size of the window
window_width = 500
window_height = 550
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Change the background color of the window
root.configure(background="lightblue")

# Start Date Frame
start_frame = tk.Frame(root)
start_frame.pack(pady=10)

# Start Date Label
start_label = tk.Label(start_frame, text="Start\nDate:", bg="black", fg="white")
start_label.pack(side="left")

# Start Date Calendar
start_cal = Calendar(start_frame, selectmode="day", date_pattern="yyyy-mm-dd")
start_cal.pack(side="left")

# End Date Frame
end_frame = tk.Frame(root)
end_frame.pack(pady=10)

# End Date Label
end_label = tk.Label(end_frame, text="End\nDate:", bg="black", fg="white")
end_label.pack(side="left")

# End Date Calendar
end_cal = Calendar(end_frame, selectmode="day", date_pattern="yyyy-mm-dd")
end_cal.pack(side="left")

# Submit Button
submit_button = tk.Button(root, text="Submit", command=get_data_for_these_dates)
submit_button.pack(pady=10)

# Output Label
output_label = tk.Label(root, 
                        text="Note: Please find the generated Excel output in this folder itself.",
                        font=("Times", 12, "bold"),
                        bg="red", fg="yellow")
                        # font=font.Font(weight="bold"))
output_label.pack()

# Create the footer label
footer_label = tk.Label(root, text="Version 1.3 (2nd August 2023)", relief=tk.RAISED, anchor=tk.W)
footer_label.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()
