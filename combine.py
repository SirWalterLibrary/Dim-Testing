from tkinter import filedialog
import pandas as pd

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
        
        # Filter out the dims with Status 3 = 00000000 (future will have the option to check which filters you want)
        # filtered_df = combined_df[combined_df["Status 3"] != 00000000]

    # Write to a combined log file
    combined_df.to_csv('Old_APU.log', index=False, sep=';')
else:
    print("No files selected")
