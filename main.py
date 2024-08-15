import os
import sys
import joblib
import pandas as pd
from tkinter import filedialog

def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

# To load the model
knn = joblib.load(resource_path('model.joblib'))

# Load the validation dataset
validate = pd.read_csv(resource_path('Xactual.csv'))

# Select the log file
log_file = filedialog.askopenfilename(
    title="Select a Log File",
    filetypes=[("Log Files", "*.log"), ("All Files", "*.*")]
)

# Print the name of the path of log file
print("Current file path:", os.path.abspath(log_file), "\n")

# Create a dataframe from the log file
X_meas = pd.read_csv(log_file, sep=';')  # new box dimensions
X_meas = X_meas[X_meas["Status 3"] != 00000000] # drops all zero dimensions
X_meas[['Length', 'Width', 'Height']] = (X_meas[['Length', 'Width', 'Height']]).round(1) # rounds all L, W, H to nearest tenth (5.799999 -> 5.8)

# Predicts the actual dimensions based on measured data
X_meas.loc[:,'Box'] = knn.predict((X_meas[['Length', 'Width', 'Height']]))

# Merge the DataFrames based on the label column, preserving the order of X_meas
merged_df = X_meas.merge(validate, on='Box', suffixes=(' Measured', ' Actual'), how='left')

# Rename the columns
merged_df.rename(columns={
    'Length Measured': 'Length',
    'Width Measured': 'Width',
    'Height Measured': 'Height'
}, inplace=True)

# Subtract the corresponding features
for col in ['Length', 'Width', 'Height']:
    merged_df[f'Δ{col}'] = merged_df[f'{col}'] - merged_df[f'{col} Actual']

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
    print(f"Height is off:  {count_ohi} out of {total_rows} time(s)")

else:
    print(f"No errors!") 

# Filter the dimensions that failed
mask = (round(merged_df['ΔLength'].abs(), 1) > 0.2) | \
       (round(merged_df['ΔWidth'].abs(), 1) > 0.2) | \
       (round(merged_df['ΔHeight'].abs(), 1) > 0.2)

filtered_df = merged_df[mask]
total_bad = filtered_df.shape[0]

# Count occurrences of each unique set of box have failed
failure_counts = filtered_df.groupby(['Box']).size()
failed_boxes = filtered_df[['Index', 'Length', 'Width', 'Height', 'Box']].to_string(index=False)

# Print boxes that fail
for label, count in failure_counts.items():
    print(f"\nBox {label} is out of spec {count} time(s)")

print(f"\n{total_bad} out of {total_rows} boxes failed:\n", failed_boxes)

# Get the base name (filename with extension)
base_name = os.path.basename(log_file)

# Split the base name into name and extension
file_name, file_extension = os.path.splitext(base_name)
output_df = merged_df[['Index', 'Length', 'Width', 'Height', 'Box', 'ΔLength', 'ΔWidth', 'ΔHeight']]


# Create a directory to store HTML files
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)

output_df.to_excel(os.path.join(output_dir, file_name + '.xlsx'), index=False)