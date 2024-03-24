import tkinter as tk
from tkinter import messagebox
import random
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# File paths
workers_file = 'workers_data.json'
assigned_file = 'assigned_data.xlsx'

# Check if workers data file exists
try:
    with open(workers_file, 'r') as file:
        workers = json.load(file)
except FileNotFoundError:
    # Default data if file doesn't exist
    workers = {
        'Day Shift': {
            'Alice': ['Pak 5 (Truck)', 'Pak 5 (Gulv)', 'Pak 1', 'Pak 2', 'Pak 3', 'Pak 4', 'Mezzanin', 'Truck Job',
                      'Make 301', 'Make 302',
                      'Make 303'],
            'Bob': ['Pak 5 (Truck)', 'Pak 5 (Gulv)', 'Pak 1', 'Pak 2', 'Pak 3', 'Pak 4', 'Mezzanin', 'Truck Job',
                    'Make 301', 'Make 302',
                    'Make 303'],
            'Emily': ['Pak 5 (Truck)', 'Pak 5 (Gulv)', 'Pak 1', 'Pak 2', 'Pak 3', 'Pak 4', 'Mezzanin', 'Truck Job',
                      'Make 301', 'Make 302',
                      'Make 303']
        },
        'Evening Shift': {
            'Charlie': ['Make 301', 'Make 302', 'Make 303'],
            'David': ['Make 301', 'Make 302', 'Make 303']
        }
    }

# Sample data
stations = ['Pak 5 (Truck)', 'Pak 5 (Gulv)', 'Pak 1', 'Pak 2', 'Pak 3', 'Pak 4', 'Mezzanin', 'Truck Job', 'Make 301',
            'Make 302', 'Make 303']
shifts = ['Day Shift (6-14)', 'Evening Shift (14-22:30)']
intervals_day_shift = ['6-10', '10-14', '14-18', '18-22:30']
intervals_evening_shift = ['14-18', '18-22:30']

# Create the main window
window = tk.Tk()
window.title("ALFA LAVAL DC - Create Rotationsplan")

# Add padding around the elements
pad_x = 20
pad_y = 10

# Label for workers list
tk.Label(window, text="Select Workers").grid(row=0, column=0, columnspan=2, padx=pad_x, pady=pad_y)

# Display the list with padding
workers_listbox = tk.Listbox(window, selectmode=tk.MULTIPLE, width=40, height=8)

# Populate workers in the listbox
for shift in workers:
    for worker in workers[shift]:
        workers_listbox.insert(tk.END, f"{worker} ({shift})")

workers_listbox.grid(row=1, column=0, columnspan=2, padx=pad_x, pady=pad_y)

# Static stations display with colors
for i, station in enumerate(stations):
    if i < len(stations) // 2:
        bg_color = 'lightyellow' if i % 2 == 0 else 'lightcyan'  # Alternating background colors
        tk.Label(window, text=station, bg=bg_color).grid(row=i + 1, column=2, padx=pad_x, pady=pad_y)
    else:
        j = i - len(stations) // 2
        bg_color = 'lightyellow' if j % 2 == 0 else 'lightcyan'  # Alternating background colors
        tk.Label(window, text=station, bg=bg_color).grid(row=j + 1, column=3, padx=pad_x, pady=pad_y)

# Create a StringVar to update the assigned workers dynamically
assigned_workers_var = tk.StringVar()


# Function to assign workers to stations randomly
def assign_workers():
    selected_workers_indices = workers_listbox.curselection()

    if not selected_workers_indices:
        messagebox.showerror("Error", "Please select workers.")
        return

    selected_workers = [workers_listbox.get(index) for index in selected_workers_indices]
    assigned_workers_list = []
    assigned_stations_intervals = {interval: [] for interval in intervals_day_shift + intervals_evening_shift}

    for worker_info in selected_workers:
        worker_name, shift = worker_info.split(" (")
        shift = shift.rstrip(")")

        if shift not in workers or worker_name not in workers[shift]:
            continue

        intervals = intervals_day_shift if shift == 'Day Shift' else intervals_evening_shift

        for interval in intervals:
            available_stations = [s for s in workers[shift][worker_name] if s not in assigned_stations_intervals[interval]]
            if available_stations:
                station = random.choice(available_stations)
                assigned_stations_intervals[interval].append(station)
                assigned_workers_list.append((worker_name, station, interval))

    team_manager_name = tm_name_entry.get()
    save_assigned_data(assigned_workers_list, team_manager_name)


# Button to trigger assignment with a colored background
assign_button = tk.Button(window, text="Assign Workers", command=assign_workers, bg='lightgrey')
assign_button.grid(row=2, column=0, columnspan=2, padx=pad_x, pady=pad_y)

# Label to display assigned workers
assigned_label = tk.Label(window, textvariable=assigned_workers_var)
assigned_label.grid(row=3, column=0, columnspan=2, padx=pad_x, pady=pad_y)

# Add a label and entry field for Team Manager's name
tm_name_label = tk.Label(window, text="Team Manager:")
tm_name_label.grid(row=4, column=0, padx=pad_x, pady=pad_y)
tm_name_entry = tk.Entry(window)
tm_name_entry.grid(row=4, column=1, padx=pad_x, pady=pad_y)


# Function to save assigned data to Excel file with stations, time intervals, and workers' names
def save_assigned_data(data_list, team_manager_name):
    wb = Workbook()
    ws = wb.active

    ws.append(['Team Manager', team_manager_name])

    stations_row = ['', 'Stations'] + intervals_day_shift + intervals_evening_shift
    ws.append(stations_row)

    assigned_workers_dict = {interval: {station: [] for station in stations} for interval in
                             intervals_day_shift + intervals_evening_shift}

    # Populate the assigned workers dictionary
    for worker, station, interval in data_list:
        if interval in assigned_workers_dict and station in assigned_workers_dict[interval]:
            assigned_workers_dict[interval][station].append(worker)
        else:
            messagebox.showerror("Error", f"Invalid station or interval: {station}, {interval}")

    # Iterate through each station and create rows for each station
    for station in stations:
        row_data = [''] * (len(intervals_day_shift) + len(intervals_evening_shift) + 2)  # Initialize row with empty values
        row_data[1] = station  # Set the station name in the second column

        # Iterate through each time interval
        for interval in intervals_day_shift + intervals_evening_shift:
            if interval in assigned_workers_dict and station in assigned_workers_dict[interval]:
                workers_assigned = assigned_workers_dict[interval][station]  # Get assigned workers for this interval and station
                row_data[intervals_day_shift.index(interval) + 2] = ', '.join(workers_assigned)  # Concatenate workers' names with commas
            else:
                row_data[intervals_day_shift.index(interval) + 2] = ''  # Leave the cell empty if no workers are assigned

        ws.append(row_data)  # Add the completed row_data to the worksheet

    # Auto-fit columns after adding data
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width

    # Save the workbook
    assigned_file = 'assigned_data.xlsx'
    wb.save(assigned_file)
    messagebox.showinfo("Save", "Assigned data saved to Excel successfully.")


# Save button to save assigned data to Excel file
save_to_excel_button = tk.Button(window, text="Save to Excel", command=lambda: save_assigned_data(assigned_workers_var.get().split('\n'), tm_name_entry.get()), bg='lightgrey')
save_to_excel_button.grid(row=5, column=0, columnspan=2, padx=pad_x, pady=pad_y)

# Run the main loop
window.mainloop()
