import pandas as pd
import re
import os

# Prompt the user to input the file path
# While entering the file path make sure not to use '' or ""
file_path = input("Enter the file path of the Excel file: ")

# Check if the file path ends with '.xls' or '.xlsx' extension
if not (file_path.lower().endswith('.xls') or file_path.lower().endswith('.xlsx')):
    print("Please provide a valid Excel file (.xls or .xlsx)")
    exit()

# Check if the file exists
if not os.path.isfile(file_path):
    print("File not found.")
    exit()

# Load the Excel file into a pandas DataFrame
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print("File not found.")
    exit()

# Search for the pattern in the DataFrame
pattern = re.compile("Not Available", flags=re.IGNORECASE)
matches = df.applymap(lambda cell: bool(pattern.search(str(cell))))

# Filter the DataFrame to include only rows with the desired pattern
filtered_df = df[matches.any(axis=1)]

# Check if the pattern is found in any cell
if not filtered_df.empty:
    print('Found in the Excel file.')
    # If you want to print the positions of the occurrences, you can use:
    print('Positions:', list(zip(*matches[matches == True].stack().index)))

    # Save the filtered DataFrame to a new Excel file
    # naming the new output file path with an xlsx format is a must
    output_file_path = input("Enter the output file path to save the results: ")
    filtered_df.to_excel(output_file_path, index=False)
    print(f'Results saved to {output_file_path}')
else:
    print("Not Found")
