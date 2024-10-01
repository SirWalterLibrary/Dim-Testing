import os
import sys
import json
import time
import joblib
import logging
import pandas as pd
import tkinter as tk 
from tkinter import filedialog, ttk 
from itertools import permutations

def validate_numeric_input(action, value_if_allowed):
    # Validate numeric input (allows decimal values)
    if action != '1':  # If action is not '1', it means it's not an insert action
        return True
    try:
        float(value_if_allowed)
        return True
    except ValueError:
        return False

def save_excel_file(output_df, excel_file, stdout):
    # Reset stdout back to original
    sys.stdout = stdout  

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
    file_path = filedialog.askopenfilename(filetypes=[("Log Files", "*.log"), ("All Files", "*.*")])
    if file_path:
        # Extract the last few parts of the path
        truncated_path = os.path.join(*file_path.split(os.sep)[-6:])  # Adjust the number as needed to control how much of the path to show
        log_file_entry.delete(0, "end")  # Deletes entry if present
        log_file_entry.insert(0, truncated_path)  # Inserts the entry with the truncated path
        return file_path
    else:
        log_file_entry.config(text="No file selected.")
        return None
    
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
    global selected_boxes, filtered_boxes  # Declare filtered_df as a global variable
    # Get the selected box sizes
    selected_boxes = [box for box, var in checkboxes.items() if var.get() == '1']

        # Save the selected boxes to a file
    save_selected_boxes(selected_boxes)
    
    # Filter the DataFrame to include only the selected box sizes
    filtered_boxes = df[df['Box'].isin(selected_boxes)]
        
    # Close the UI
    window.destroy()

def setup_logging(folder):
    # Only set up logging if an error is detected
    log_file = os.path.join(folder, "error.log")
    logging.basicConfig(filename=log_file, level=logging.ERROR)

    return logging

def main():
    global log_file_entry, box_df, length_tol_entry, width_tol_entry, height_tol_entry, checkboxes

    # Create the main window
    root = tk.Tk()
    root.title("Log Parser")
    
    # Make the window stay on top of other windows
    root.attributes('-topmost', True)

    # Set a fixed window size
    root.geometry("500x400")  # You can adjust the size as needed

    # Validation to ensure only numeric values (including decimals) are entered
    vcmd = (root.register(validate_numeric_input), '%d', '%P')

    # Create a frame for importing a log file
    frame_import = tk.Frame(root, padx=10, pady=10)
    frame_import.grid(row=0, column=0, columnspan=3, sticky="ew")

    import_button = tk.Button(frame_import, text="Import Log File", command=load_file)
    import_button.grid(row=0, column=0, padx=5)

    log_file_entry = tk.Entry(frame_import, text="No file selected.", width=55)
    log_file_entry.grid(row=0, column=1, padx=10, sticky="w")

    # Create a main frame to hold tolerances and boxes
    frame_main = tk.Frame(root, padx=10, pady=10)
    frame_main.grid(row=1, column=0, columnspan=3, sticky="ew")

    # Create sub-frame for tolerances (top side)
    frame_tolerances = tk.Frame(frame_main)
    frame_tolerances.grid(row=0, column=0, columnspan=3, sticky="ew")

    tk.Label(frame_tolerances, text="Enter Tolerances:").grid(row=0, column=0, sticky=tk.W)

    # Length tolerance
    tk.Label(frame_tolerances, text="Length Tolerance:").grid(row=1, column=0, sticky=tk.W)
    length_tol_entry = tk.Entry(frame_tolerances, width=5, validate="key", validatecommand=vcmd)
    length_tol_entry.grid(row=1, column=1, pady=5)

    # Width tolerance
    tk.Label(frame_tolerances, text="Width Tolerance:").grid(row=2, column=0, sticky=tk.W)
    width_tol_entry = tk.Entry(frame_tolerances, width=5, validate="key", validatecommand=vcmd)
    width_tol_entry.grid(row=2, column=1, pady=5)

    # Height tolerance
    tk.Label(frame_tolerances, text="Height Tolerance:").grid(row=3, column=0, sticky=tk.W)
    height_tol_entry = tk.Entry(frame_tolerances, width=5, validate="key", validatecommand=vcmd)
    height_tol_entry.grid(row=3, column=1, pady=5)

    # Run the Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    # Run the main program
    main()