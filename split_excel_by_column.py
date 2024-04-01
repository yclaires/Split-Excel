""" This module provide functionality to split excel files. 
"""
import argparse
from pathlib import Path
import os
import pandas as pd

def split_excel(excel_path:Path, excel_column:str):
    
    # Read the Excel file
    # file_path = 'Excel-Split.xlsx'  # Replace 'input_file.xlsx' with your Excel file path
    df = pd.read_excel(excel_path)

    # Identify unique values in the column for splitting
    # excel_column = 'Branch'  # Replace 'Column_Name' with the name of the column you want to split by
    unique_values = df[excel_column].unique()

    # Create a target directory to save the split Excel files
    target_directory = Path.cwd() / 'output_folder'  # Replace 'output_folder' with the desired target directory name
    #target_directory = Path.cwd().joinpath('output_folder')
    target_directory.mkdir(exist_ok=True)
    #os.makedirs(target_directory, exist_ok=True) # Use os.makedirs function to create the target directory

    # Split the data based on unique values in the specified column
    for value in unique_values:
        # Create a subset of the DataFrame based on the current unique value
        subset = df[df[excel_column] == value]
    
        # Define the output file path
        #output_file_path = os.path.join(target_directory, f'{value}.xlsx')  # Use 'os.path.join()' function to define the output file
        output_file_path = Path(target_directory) / f'{value}.xlsx'  # Use ' Path() / ' to define the output file

    
        # Write the subset to a new Excel file
        subset.to_excel(output_file_path, index=False)
        print(f"Excel file saved: {output_file_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description = 'split excel by column value')
    parser.add_argument('file_path', metavar = 'file_parth', type = str, help = 'enter your file path')
    parser.add_argument('split_column', metavar = 'split_column', type = str, help = 'enter your split column name')
    args = parser.parse_args()

    file_path = args.file_path
    split_column = args.split_column
    split_excel(file_path, split_column)