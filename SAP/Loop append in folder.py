import os
import pandas as pd

# Folder containing Excel files
folder_path = r"D:\My\Project\KPI_mspha_automate\Destination_data"

# Check if folder exists
if not os.path.exists(folder_path):
    print(f"Error: Folder not found -> {folder_path}")
    exit()

# List all Excel files in the folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith(".XLSX")]

# Check if there are files
if not excel_files:
    print("No .xlsx files found in the folder!")
    exit()

# Initialize an empty list to store DataFrames
all_data = []

# Loop through each file and append data
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    print(f"Processing file: {file_path}")

    try:
        df = pd.read_excel(file_path, engine="openpyxl")

        if df.empty:
            print(f"Warning: {file} is empty!")
            continue  # Skip empty files

        df["Source_File"] = file  # Add a column to track the source file
        all_data.append(df)  # Append DataFrame to list

        print(f"Appended {len(df)} rows from {file}")

    except Exception as e:
        print(f"Error reading {file}: {e}")

# Combine all DataFrames into one
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    print("All files processed successfully!")
else:
    print("No valid data found to append!")
