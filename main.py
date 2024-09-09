import pandas as pd
import json

# Replace with the path to your JSON file
json_file_path = r'C:\Users\Froap\_DEV\Planner_Tasks_to_Excel\TASKS.json'

# Load JSON data from the file
with open(json_file_path, 'r') as file:
    json_data = json.load(file)

# Extract task details
tasks = json_data['value']

# Normalize JSON data to create a DataFrame
df_tasks = pd.json_normalize(tasks)

# Save DataFrame to an Excel file
df_tasks.to_excel('planner_tasks.xlsx', index=False)

print("JSON data has been successfully converted to an Excel file.")
