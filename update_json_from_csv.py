import json
import csv

# Load JSON data from workers_data.json
with open('workers_data.json', 'r') as json_file:
    workers_data = json.load(json_file)

# Read CSV data and update JSON
with open('assigned_data.csv', 'r') as csv_file:
    csv_reader = csv.DictReader(csv_file)
    for row in csv_reader:
        worker_id = row['worker_id']  # Assuming 'worker_id' is a unique identifier
        if worker_id in workers_data:
            # Update the corresponding fields in JSON
            workers_data[worker_id]['name'] = row['name']
            workers_data[worker_id]['age'] = int(row['age'])  # Convert to integer if needed
            # Update other fields as necessary

# Write the updated JSON back to workers_data.json
with open('workers_data.json', 'w') as json_file:
    json.dump(workers_data, json_file, indent=4)
