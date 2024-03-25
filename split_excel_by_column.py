import pandas as pd
import os

# Read the Excel file
file_path = 'Excel-Split.xlsx'  # Replace 'input_file.xlsx' with your Excel file path
df = pd.read_excel(file_path)

# Identify unique values in the column for splitting
split_column = 'Kontinent'  # Replace 'Column_Name' with the name of the column you want to split by
unique_values = df['Kontinent'].unique()

# Create a target directory to save the split Excel files
target_directory = 'output_folder'  # Replace 'output_folder' with the desired target directory name
os.makedirs(target_directory, exist_ok=True)

# Split the data based on unique values in the specified column
for value in unique_values:
    # Create a subset of the DataFrame based on the current unique value
    subset = df[df['Kontinent'] == value]
    
    # Define the output file path
    output_file_path = os.path.join(target_directory, f'{value}.xlsx')
    
    # Write the subset to a new Excel file
    subset.to_excel(output_file_path, index=False)
    print(f"Excel file saved: {output_file_path}")
