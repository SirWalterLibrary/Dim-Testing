import os
import sys
import json
import time
import joblib
import logging
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from itertools import permutations

def save_excel_file(output_df, excel_file):
    # Extract the file name without the folder name
    excel_filename = os.path.basename(excel_file)

    while True:
        try:
            # Try to save the DataFrame to Excel
            output_df.to_excel(excel_file, index=False)
            break  # Exit the loop once the file is saved
        except PermissionError:
            # This exception occurs if the file is open in Excel
            print(f"\nUnable to save results to an Excel file because \"{excel_filename}\" is currently open. Please close the file to proceed.")
            user_input = input("Are you ready to retry? (Y/n): ").strip().lower()
            if user_input == 'n':
                print("Exiting without saving.")
                break
            elif user_input == 'y':
                print("Retrying to save the file...")
                time.sleep(1)  # Wait for a second before retrying
            else:
                print("Invalid input. Please type 'Y' or 'n'.")

def load_file():
    while True:
        try:
            # Open file dialog to select a log file
            file = filedialog.askopenfilename(
                title="Select a Log File",
                filetypes=[("Log Files", "*.log"), ("All Files", "*.*")]
            )
            
            # If user cancels the dialog, file will be empty
            if not file:
                print("No file selected. Exiting...")
                return None
            
            return file

        except PermissionError:
            # Handle the case when the file is open in Excel or locked
            print("PermissionError: The log file is open or locked.")
        
        # Prompt the user if they want to retry
        while True:
            user_input = input("Do you want to retry? (Y/n): ").strip().lower()
            if user_input == 'n':
                print("Exiting...")
                return None
            elif user_input == 'y':
                print("Retrying...")
                time.sleep(1)  # Optional sleep before retry
                break
            else:
                print("Invalid input. Please type 'Y' or 'n'.")

def calculate_min_difference(row):
    # Specify Actual vs Result dimensions
    actual_dims = (row['Length Actual'], row['Width Actual'], row['Height Actual'])
    result_dims = (row['Length'], row['Width'], row['Height'])
    
    # Generate all permutations (rotations) of the expected dimensions
    rotations = list(permutations(actual_dims))
    
    # Initialize parameters for finding the best rotation
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

def save_selected_boxes(selected_boxes, file_path=resource_path("selected_boxes.json")):
    """Save selected boxes to a file."""
    with open(file_path, 'w') as f:
        json.dump(selected_boxes, f)

def load_selected_boxes(file_path=resource_path("selected_boxes.json")):
    """Load selected boxes from a file."""
    if os.path.exists(file_path):
        with open(file_path, 'r') as f:
            return json.load(f)
    return None

def filter_boxes(window, df, checkboxes):
    global filtered_boxes  # Declare filtered_df as a global variable
    # Get the selected box sizes
    selected_boxes = [box for box, var in checkboxes.items() if var.get() == '1']

        # Save the selected boxes to a file
    save_selected_boxes(selected_boxes)
    
    # Filter the DataFrame to include only the selected box sizes
    filtered_boxes = df[df['Box'].isin(selected_boxes)]
        
    # Close the UI
    window.destroy()

# Get the user's Downloads folder path
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

def main():
    # Redirect stdout to a file
    try:
        # Define the output directory and filename
        dims_file = load_file()
        if not dims_file:
            return

        # Get the base name (filename with extension)
        base_name = os.path.basename(dims_file)

        # Split the base name into filename and extension
        file_name, file_extension = os.path.splitext(base_name)

        # Create a directory with filename
        output_path = os.path.join(downloads_folder, f"output/{file_name}")
        os.makedirs(output_path, exist_ok=True)

        summary_file = os.path.join(output_path, "summary.txt")

        # Redirect stdout to the summary file
        with open(summary_file, 'w') as file:
            original_stdout = sys.stdout  # Save the original stdout
            sys.stdout = file  # Redirect stdout to the file

            # To load the model
            knn = joblib.load(resource_path('model.joblib'))

            # Load the CSV with actual dimensions
            box_df = pd.read_csv(resource_path('Xactual.csv'))

            # Get unique box sizes from the "Box" column
            unique_boxes = box_df['Box'].unique()

            # Load previously saved selected boxes (if available)
            previous_selected_boxes = load_selected_boxes()

            # Create the tkinter window
            root = tk.Tk()
            root.title("Select Relevant Box Sizes")

            # Make the window stay on top of other windows
            root.attributes('-topmost', True)

            # Create a dictionary to hold checkboxes
            checkboxes = {}

            # Add a checkbox for each unique box size
            for box in unique_boxes:
                var = tk.StringVar(value='1' if previous_selected_boxes and box in previous_selected_boxes else '0')  # Set default based on saved data
                checkbox = ttk.Checkbutton(root, text=box, variable=var)
                checkbox.pack(anchor='w')
                checkboxes[box] = var

            # Add a button to confirm selection
            button = ttk.Button(root, text="Filter Boxes", command=lambda: filter_boxes(root, box_df, checkboxes))
            button.pack()

            # Run the tkinter main loop
            root.mainloop()

            # Create a dataframe of all the measurements from the log file
            meas_df = pd.read_csv(dims_file, sep=';')  # new box dimensions
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

            # Align Actual vs Result dimensions (i.e. "5.2x6.2x2.0" would be "6x5x2" box but calculated difference would be "5x6x2")
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
                print(f"All populated dimensions are within spec!") 

            # Filter the dimensions that failed
            mask = (round(merged_df['ΔLength'].abs(), 1) > 0.2) | \
                (round(merged_df['ΔWidth'].abs(), 1) > 0.2) | \
                (round(merged_df['ΔHeight'].abs(), 1) > 0.2)

            # Filter out the bad dimensions and count number of occurrences
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

            # Calculate and print the success rate; print failed boxes if applicable
            success_rate = (total_rows - total_bad) / total_rows * 100
            print(f"\n{total_bad} out of {total_rows} boxes failed: {success_rate:.2f}% success rate\n", "\nFailed boxes:\n", failed_boxes if total_bad else "")



            # Create an output DataFrame to export to an Excel file with only valid dimensions
            output_df = merged_df[merged_df["Box"] != "0x0x0"][['Index', 'Length', 'Width', 'Height', 'Box', 'ΔLength', 'ΔWidth', 'ΔHeight']]

            # Set the full path for the Excel file
            excel_file = os.path.join(output_path, file_name + '.xlsx')

            # Call the save function with the DataFrame and file path
            time.sleep(1)
            save_excel_file(output_df, excel_file)

            # Optionally, open the saved file
            os.startfile(output_path)

        sys.stdout = original_stdout  # Reset stdout back to original

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Run the main program
    main()