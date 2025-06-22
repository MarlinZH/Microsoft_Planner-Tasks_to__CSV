import pandas as pd
import json
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Hide the root Tkinter window
Tk().withdraw()

# Open a file dialog for the user to select the JSON file
json_file_path = askopenfilename(
    title="Select your Planner Tasks JSON file",
    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
)

# Load JSON data from the file
if not json_file_path:
    print("No file selected. Exiting.")
    exit()
    
with open(json_file_path, 'r') as file:
    json_data = json.load(file)

# Extract task details
tasks = json_data['value']

# Normalize JSON data to create a DataFrame
df_tasks = pd.json_normalize(tasks)

# Save DataFrame to an Excel file
df_tasks.to_excel('planner_tasks.xlsx', index=False)

print("JSON data has been successfully converted to an Excel file.")
