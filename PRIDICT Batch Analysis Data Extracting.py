import os
import csv
from openpyxl import load_workbook

# Directory containing the CSV files
directory = "/Users/qichenyuan/PRIDICT/batch"

# Create a list to store unique names
unique_names = set()

# Iterate over files in the directory
for filename in os.listdir(directory):
    if "nicking_guides.csv" in filename or "Pridict_full.csv" in filename:
        # Extract the name part
        name_part = filename.split("_")[0]
        unique_names.add(name_part)

# Open the Excel file
excel_file_path = "/Users/qichenyuan/PRIDICT/batch/FINAL List NAFLD.xlsx"
wb = load_workbook(excel_file_path)
ws = wb.active

# Write unique names to Sheet1 Column A
ws["A1"] = "sequence_name"
for i, name in enumerate(unique_names, start=2):
    ws[f"A{i}"] = name

# Save the changes
wb.save(excel_file_path)


# Path to Excel file
excel_file_path = "/Users/qichenyuan/PRIDICT/batch/FINAL List NAFLD.xlsx"

# Path to CSV folder
csv_folder_path = "/Users/qichenyuan/PRIDICT/batch"

# Load Excel workbook
workbook = load_workbook(excel_file_path)
sheet = workbook.active

# Iterate over rows in Excel sheet, starting from the second row
for row in sheet.iter_rows(min_row=2, max_col=1, max_row=sheet.max_row):
    sequence_name = row[0].value
    csv_file_path = os.path.join(csv_folder_path, f"{sequence_name}_nicking_guides.csv")

    # Check if CSV file exists
    if os.path.exists(csv_file_path):
        # Open and read CSV file
        with open(csv_file_path, 'r') as csv_file:
            lines = csv_file.readlines()
            # Extract B2 value
            if len(lines) > 1:
                b2_value = lines[1].split(',')[1].strip()  # Assuming B2 means row 2, column 2 (0-indexed)

                # Get the cell in column N corresponding to the current row
                cell = sheet.cell(row=row[0].row, column=14)  # Column N is the 14th column

                # Set the value of the cell in column N to the extracted value
                cell.value = b2_value

# Save updated Excel workbook
workbook.save(excel_file_path)


# Define paths
excel_file_path = "/Users/qichenyuan/PRIDICT/batch/FINAL List NAFLD.xlsx"
csv_folder_path = "/Users/qichenyuan/PRIDICT/batch"

# Load Excel workbook
workbook = load_workbook(excel_file_path)
sheet = workbook.active

# Find the first empty row after row 1
empty_row = 2
while sheet.cell(row=empty_row, column=1).value is not None:
    empty_row += 1

# Start filling data from B2
current_row = 2

# Iterate through each row in the Excel file
for row in sheet.iter_rows(min_row=current_row, max_col=1, max_row=empty_row - 1, values_only=True):
    sequence_name = row[0]

    # Construct the CSV file path
    csv_file_path = os.path.join(csv_folder_path, f"{sequence_name}_pegRNA_Pridict_full.csv")

    # Check if CSV file exists
    if os.path.exists(csv_file_path):
        # Read the CSV file
        with open(csv_file_path, 'r') as csv_file:
            csv_reader = csv.reader(csv_file)
            csv_data = list(csv_reader)

            # Extract values from specific rows and columns
            b_value = csv_data[1][1]
            c_value = csv_data[1][2]
            n_value = csv_data[1][13]
            q_value = csv_data[1][16]
            o_value = csv_data[1][14]

            # Update the Excel file with extracted values
            sheet.cell(row=current_row, column=2, value=b_value)
            sheet.cell(row=current_row, column=3, value=c_value)
            sheet.cell(row=current_row, column=4, value=n_value)
            sheet.cell(row=current_row, column=6, value=q_value)
            sheet.cell(row=current_row, column=7, value=o_value)

            # Move to the next row
            current_row += 1

# Save the updated Excel file
workbook.save(excel_file_path)


