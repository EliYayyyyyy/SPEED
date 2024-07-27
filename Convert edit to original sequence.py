import pandas as pd
import re

# Read the Excel file
input_file_path = '/Users/eliyayyyyyy/Desktop/BootStrap Bio/Previous /ANGEL.xlsx'
sheet_name = 'Sheet1'

# Load the Excel file into a DataFrame
df = pd.read_excel(input_file_path, sheet_name=sheet_name)

# Check if the 'Edit Sequence' column exists
if 'Edit Sequence' not in df.columns:
    raise ValueError("Column 'Edit Sequence' not found in the DataFrame. Please check the column name.")

# Define a function to process DNA sequences
def process_sequence(sequence):
    # Substitute (X/Y) with X
    sequence = re.sub(r'\(([^/)]+)/([^)]+)\)', r'\1', sequence)
    # Remove (+X)
    sequence = re.sub(r'\(\+([^)]+)\)', '', sequence)
    # Replace (-X) with X
    sequence = re.sub(r'\(-([^)]+)\)', r'\1', sequence)
    return sequence

# Apply the function to each sequence in the DataFrame and store the result in a new column
df['Original Sequence'] = df['Edit Sequence'].apply(process_sequence)

# Save the modified DataFrame back to the Excel file
output_file_path = '/Users/eliyayyyyyy/Desktop/BootStrap Bio/Previous /ANGEL.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Processed sequences saved to {output_file_path}")


