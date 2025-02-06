from tkinter import filedialog
import pandas as pd
import os

# Open file dialog to select log files
file_paths = filedialog.askopenfilenames(title="Select Log Files", filetypes=[("Log Files", "*.log")])

if file_paths:
    # Initialize an empty list to store DataFrames
    df_list = []
    
    # Read each file and extract data
    for file_path in file_paths:
        df = pd.read_csv(file_path, sep=';')
        df_list.append(df)
    
    # Concatenate all DataFrames into a single DataFrame
    combined_df = pd.concat(df_list, ignore_index=True)
    
    # Determine the output directory
    directories = [os.path.dirname(file_path) for file_path in file_paths]
    if len(set(directories)) == 1:
        output_dir = directories[0]
    else:
        output_dir = os.path.join(os.path.dirname(__file__), 'output')
        os.makedirs(output_dir, exist_ok=True)
    
    # Write to a combined log file
    output_path = os.path.join(output_dir, 'dims.log')
    combined_df.to_csv(output_path, index=False, sep=';')
    print(f"Combined log file saved to: {output_path}")
else:
    print("No files selected")
