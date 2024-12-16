

import joblib
import re
import pandas as pd
import math
import os


# Dictionary to store the latest instance of each unique ID
id_dict = {}

# Read the file
with open(r"C:\Users\mohammo\Documents\Projects\UPS 5200\VMSSorter_9379352_T3.31.5.0\output.txt", 'r') as file:
    log_data = file.readlines()

# Regex patterns for "before correction" and "after correction"
pattern_before = re.compile(r"VMS Result before correction:\s+ID\s*:\s*(\d+),\s*L\s*:\s*(\d+),\s*W\s*:\s*(\d+),\s*H\s*:\s*(\d+)")
pattern_after = re.compile(r"VMS Result after correction:\s+ID\s*:\s*(\d+),\s*L\s*:\s*(\d+),\s*W\s*:\s*(\d+),\s*H\s*:\s*(\d+)")

# Iterate through the lines and find matches for both "before" and "after"
for line in log_data:
    # Check for "before correction" result
    match_before = pattern_before.search(line)
    if match_before:
        id_value = int(match_before.group(1))
        length = int(match_before.group(2))
        width = int(match_before.group(3))
        height = int(match_before.group(4))
        
        # Convert from mm to inches
        length_in_inches = length / 2.54
        width_in_inches = width / 2.54
        height_in_inches = height / 2.54
        
        # Apply the rounding formula
        length_in_inches_rounded = math.floor((length_in_inches / 20) + 0.5) * 20 / 100
        width_in_inches_rounded = math.floor((width_in_inches / 20) + 0.5) * 20 / 100
        height_in_inches_rounded = math.floor((height_in_inches / 20) + 0.5) * 20 / 100
        
        # Store the "before correction" values in the dictionary
        id_dict[id_value] = [id_value, length_in_inches_rounded, width_in_inches_rounded, height_in_inches_rounded]
    
    # Check for "after correction" result and replace the existing entry for the same ID
    match_after = pattern_after.search(line)
    if match_after:
        id_value = int(match_after.group(1))
        length = int(match_after.group(2))
        width = int(match_after.group(3))
        height = int(match_after.group(4))
        
        # Convert from mm to inches
        length_in_inches = length / 2.54
        width_in_inches = width / 2.54
        height_in_inches = height / 2.54
        
        # Apply the rounding formula
        length_in_inches_rounded = math.floor((length_in_inches / 20) + 0.5) * 20 / 100
        width_in_inches_rounded = math.floor((width_in_inches / 20) + 0.5) * 20 / 100
        height_in_inches_rounded = math.floor((height_in_inches / 20) + 0.5) * 20 / 100
        
        # Replace the entry in the dictionary with the "after correction" values
        id_dict[id_value] = [id_value, length_in_inches_rounded, width_in_inches_rounded, height_in_inches_rounded]

# Convert the dictionary values to a DataFrame
df = pd.DataFrame(list(id_dict.values()), columns=['ID', 'Length', 'Width', 'Height'])

# Sort by ID in ascending order and set ID as the index
df = df.sort_values(by='ID').set_index('ID')

# Add Rim Plane 1 parameters as constant values across all rows
df['a1'] = 0.9794
df['b1'] = 0.0059
df['c1'] = 0.2016
df['d1'] = -40.9368

# Add Rim Plane 2 parameters as constant values across all rows
df['a2'] = 0.0058
df['b2'] = -0.9988
df['c2'] = 0.0492
df['d2'] = 490.7688

# To load the model
knn = joblib.load('model.joblib')
df.loc[:, 'Box'] = knn.predict(df[['Length', 'Width', 'Height']])

# Read the actual dimensions CSV
df_actual = pd.read_csv('Xactual.csv')

# Merge the DataFrames based on the 'Box' column (keeping index intact)
merged_df = df.reset_index().merge(df_actual, on='Box', suffixes=(' Measured', ' Actual'), how='left')

# After merging, set the index back to 'ID' if it was lost during the merge
merged_df.set_index('ID', inplace=True)

# Rename the columns for a cleaner look
merged_df.rename(columns={
    'Length Measured': 'Length',
    'Width Measured': 'Width',
    'Height Measured': 'Height'
}, inplace=True)

# Subtract the corresponding features
for col in ['Length', 'Width', 'Height']:
    merged_df[f'error_{col}'] = merged_df[f'{col}'] - merged_df[f'{col} Actual']

# Display the DataFrame to verify that constants were added correctly
output_df = merged_df[['a1', 'b1', 'c1', 'd1', 'a2', 'b2', 'c2', 'd2', 'Length', 'Width', 'Height', 'Box', 'error_Length', 'error_Width', 'error_Height']]

# Set the full path for the Excel file
excel_file = 'APU_train.csv'

# Check if the CSV file exists
if os.path.exists(excel_file):
    # If the file exists, append new data to it
    output_df.to_csv(excel_file, mode='a', header=False, index=True)
else:
    # If the file doesn't exist, create a new CSV and write the data
    output_df.to_csv(excel_file, index=True)

# Open the Excel file after modification
os.startfile(excel_file)
