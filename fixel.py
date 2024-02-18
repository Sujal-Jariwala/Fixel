import pandas as pd
import re

file_path = r'C:\Users\Sujal\Desktop\Mangaldeep_Pharmacy_Item_List.xls'

# Load the Excel file into a pandas DataFrame
df = pd.read_excel(file_path)

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
    output_file_path = r'C:\Users\Sujal\Desktop\Results.xlsx'
    filtered_df.to_excel(output_file_path, index=False)
    print(f'Results saved to {output_file_path}')
else:
    print("Not Found")