import os
import sys
import joblib
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from itertools import permutations

def calculate_min_difference(row):
    actual_dims = (row['Length Actual'], row['Width Actual'], row['Height Actual'])
    result_dims = (row['Length'], row['Width'], row['Height'])
    
    # Generate all permutations (rotations) of the expected dimensions
    rotations = list(permutations(actual_dims))
    
    min_difference = None
    best_rotation = None
    
    for rotation in rotations:
        # Calculate the difference for this rotation
        difference = [result_dims[i] - rotation[i] for i in range(3)]
        
        # Compute the total absolute difference
        total_difference = sum(abs(diff) for diff in difference)
        
        # Update minimum difference and best rotation
        if min_difference is None or total_difference < min_difference:
            min_difference = total_difference
            best_rotation = rotation
    
    # Format differences to .2f
    formatted_difference = [f"{diff:.2f}" for diff in [result_dims[i] - best_rotation[i] for i in range(3)]]
    
    return pd.Series(formatted_difference, index=['Difference Length', 'Difference Width', 'Difference Height'])

def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

def filter_boxes():
    global filtered_boxes  # Declare filtered_df as a global variable
    # Get the selected box sizes
    selected_boxes = [box for box, var in checkboxes.items() if var.get() == '1']
    
    # Filter the DataFrame to include only the selected box sizes
    filtered_boxes = box_df[box_df['Box'].isin(selected_boxes)]
        
    # Close the UI
    root.destroy()

# To load the model
knn = joblib.load(resource_path('model.joblib'))

# Load the CSV with actual dimensions
box_df = pd.read_csv('Xactual.csv')

# Get unique box sizes from the "Box" column
unique_boxes = box_df['Box'].unique()

# Create the tkinter window
root = tk.Tk()
root.title("Select Relevant Box Sizes")

# Make the window stay on top of other windows
root.attributes('-topmost', True)

# Create a dictionary to hold checkboxes
checkboxes = {}

# Add a checkbox for each unique box size
for box in unique_boxes:
    var = tk.StringVar(value='1')  # Set default to selected
    checkbox = ttk.Checkbutton(root, text=box, variable=var)
    checkbox.pack(anchor='w')
    checkboxes[box] = var

# Add a button to confirm selection
button = ttk.Button(root, text="Filter Boxes", command=filter_boxes)
button.pack()

# Run the tkinter main loop
root.mainloop()

# Select the log file
log_file = filedialog.askopenfilename(
    title="Select a Log File",
    filetypes=[("Log Files", "*.log"), ("All Files", "*.*")]
)

# Print the name of the path of log file
print("Current file path:", os.path.abspath(log_file), "\n")

# Create a dataframe of all the measurements from the log file
meas_df = pd.read_csv(log_file, sep=';')  # new box dimensions
meas_df = meas_df[meas_df["Status 3"] != 00000000] # drops all zero dimensions
meas_df[['Length', 'Width', 'Height']] = (meas_df[['Length', 'Width', 'Height']]).round(1) # rounds all L, W, H to nearest tenth (5.799999 -> 5.8)

# Predicts the actual dimensions based on measured data
meas_df.loc[:,'Box'] = knn.predict((meas_df[['Length', 'Width', 'Height']]))

# Merge the DataFrames based on the label column, preserving the order of meas_df
merged_df = meas_df.merge(filtered_boxes, on='Box', suffixes=(' Measured', ' Actual'), how='left')

# Rename the columns for cleaner look in Excel
merged_df.rename(columns={
    'Length Measured': 'Length',
    'Width Measured' : 'Width',
    'Height Measured': 'Height'
}, inplace=True)

# Subtract the corresponding features
for col in ['Length', 'Width', 'Height']:
    merged_df[f'Δ{col}'] = merged_df[f'{col}'] - merged_df[f'{col} Actual']

merged_df[[f'Δ{col}' for col in ['Length', 'Width', 'Height']]] = merged_df.apply(calculate_min_difference, axis=1).astype(float)

# Select only the columns containing the differences
difference_df = merged_df[[f'Δ{col}' for col in ['Length', 'Width', 'Height']]]

# Calculate the frequency each column is out of spec (greater than ± 0.2)
count_ole, count_owi, count_ohi = [sum((round(difference_df[f'Δ{dim}'].abs(), 1) > 0.2)) for dim in ['Length', 'Width', 'Height']]
    
# Calculate the number of rows with populated dimensions
total_rows = difference_df.shape[0] 

# Check whether any dimensions are out of spec
if count_ole > 0 or count_owi > 0 or count_ohi > 0:

    # Print the occurrences each time length, width, and height is off 
    print(f"Length is off:  {count_ole} out of {total_rows} time(s)")
    print(f"Width  is off:  {count_owi} ouf of {total_rows} time(s)")
    print(f"Height is off:  {count_ohi} out of {total_rows} time(s)\n")

else:
    print(f"No errors!") 

# Filter the dimensions that failed
mask = (round(merged_df['ΔLength'].abs(), 1) > 0.2) | \
       (round(merged_df['ΔWidth'].abs(), 1) > 0.2) | \
       (round(merged_df['ΔHeight'].abs(), 1) > 0.2)

filtered_df = merged_df[mask]
total_bad = filtered_df.shape[0]

# Count occurrences of each unique set of box failures
failure_counts = filtered_df.groupby(['Box']).size()

# Sort the filtered DataFrame by the 'Index' column
sorted_failed_boxes = filtered_df[['Index', 'Length', 'Width', 'Height', 'Box']].sort_values(by=['Box', 'Index'])

# Convert the sorted DataFrame to a string without the default index
failed_boxes = sorted_failed_boxes.to_string(index=False)

# Print boxes that fail
for label, count in failure_counts.items():
    print(f"Box {label} is out of spec {count} time(s)")
    
# Set the display option to expand the column width
pd.set_option('display.max_colwidth', None)

# Calculate and print the success rate
success_rate = (total_rows - total_bad) / total_rows * 100
print(f"\n{total_bad} out of {total_rows} boxes failed: {success_rate:.2f}% success rate\n\n", failed_boxes)


# Get the base name (filename with extension)
base_name = os.path.basename(log_file)

# Split the base name into name and extension
file_name, file_extension = os.path.splitext(base_name)
output_df = merged_df[['Index', 'Length', 'Width', 'Height', 'Box', 'ΔLength', 'ΔWidth', 'ΔHeight']]

# Create a directory to store HTML files
output_dir = 'excel_output'
os.makedirs(output_dir, exist_ok=True)

output_df.to_excel(os.path.join(output_dir, file_name + '.xlsx'), index=False)