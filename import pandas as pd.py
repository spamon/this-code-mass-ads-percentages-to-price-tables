import pandas as pd
import numpy as np

# Read the Excel file with multiple pricing tables
file_path = 'grouped price list.xlsx'
sheet_names = ['Sheet1']  # Modify the sheet names as per your file

# Define the percentage increase
percentage_increase = 0.10  # 10% increase, modify as needed

# Create an empty list to store the updated tables
updated_tables = []

# Iterate over each sheet
for sheet_name in sheet_names:
    # Read the pricing table from the sheet
    table = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Specify the range of cells to update
    start_row = 1  # Modify as per your requirement
    end_row = table.shape[0] - 1  # Modify as per your requirement
    start_col = 1  # Modify as per your requirement
    end_col = table.shape[1] - 1  # Modify as per your requirement
    
    # Create a copy of the table to preserve the original values
    updated_table = table.copy()
    
    # Apply the percentage increase to the specified range of cells
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_value = table.iloc[row, col]
            
            # Check if the cell value is numeric
            if isinstance(cell_value, (int, float, np.number)):
                updated_value = cell_value * (1 + percentage_increase)
                updated_table.iloc[row, col] = updated_value
    
    # Add the updated table to the list
    updated_tables.append(updated_table)

# Write the updated tables to a new Excel file
output_file = 'updated_pricing_tables.xlsx'
with pd.ExcelWriter(output_file) as writer:
    for i, sheet_name in enumerate(sheet_names):
        updated_tables[i].to_excel(writer, sheet_name=sheet_name, index=False)

print("Pricing tables updated and saved to", output_file)
