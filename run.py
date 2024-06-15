import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to select the folder containing Excel files
def select_folder():
    folder_selected = filedialog.askdirectory()
    folder_path.set(folder_selected)

# Function to select the save location for the combined Excel file
def select_save_location():
    file_selected = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    save_path.set(file_selected)

# Function to combine and save the Excel files
def combine_and_save():
    try:
        folder = folder_path.get()
        save_file = save_path.get()
        
        if not folder or not save_file:
            messagebox.showerror("Error", "Please select a folder and a save location.")
            return
        
        # Get a list of all Excel files in the folder
        excel_files = [file for file in os.listdir(folder) if file.endswith('.xlsx')]

        # Create an empty list to store the data
        combined_data = []

        # Read each Excel file and append it to the list
        for file in excel_files:
            file_path = os.path.join(folder, file)
            df = pd.read_excel(file_path)
            combined_data.append(df)

        # Combine all data into a single DataFrame
        combined_df = pd.concat(combined_data, ignore_index=True)

        # Save the combined data to a new Excel file
        combined_df.to_excel(save_file, index=False)

        messagebox.showinfo("Success", f"All Excel files have been combined into {save_file}.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the Tkinter interface
root = tk.Tk()
root.title("Combine Excel Files")

folder_path = tk.StringVar()
save_path = tk.StringVar()

tk.Label(root, text="Select the folder containing the Excel files:").pack(pady=5)
tk.Button(root, text="Select Folder", command=select_folder).pack(pady=5)
tk.Entry(root, textvariable=folder_path, width=50).pack(pady=5)

tk.Label(root, text="Select the save location for the combined file:").pack(pady=5)
tk.Button(root, text="Select Save Location", command=select_save_location).pack(pady=5)
tk.Entry(root, textvariable=save_path, width=50).pack(pady=5)

tk.Button(root, text="Combine and Save", command=combine_and_save).pack(pady=20)

root.mainloop()
