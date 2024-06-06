# Matt Gurgiolo 2024
# COOT Sorter

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from ttkthemes import ThemedStyle
from Sort import Sorter
import os

def select_file():
    global filepath
    filepath = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    if filepath:  # Check if a file was selected
        filename = os.path.basename(filepath)  # Extract the filename from the filepath
        file_path_label.config(text=filename)
    
def process_file():
    
    # Check if the filepath is set
    if not filepath:
        messagebox.showerror("Error", "No file selected. Please select a file first.")
        return
    
    try:
        # Create an instance of the Sorter class with the selected file and sort
        sorter = Sorter(filepath)
        sorter.run()

        # Save the sorted data to a new file
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[('Excel Files', '*.xlsx')])
        if save_path:
            sorter.write(save_path)
            messagebox.showinfo("Success", "File processed and saved successfully.")
        else:
            messagebox.showerror("Error", "Save file not selected.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing the file: {e}")

# The Main Window
root = tk.Tk()
root.title('COOT Sorter')
root.geometry('1000x500')

# Apply a theme
style = ThemedStyle(root)
style.configure('TFrame', background='chartreuse4')
style.configure('TLabel', background='chartreuse4', foreground='black')
style.configure('Label', background='chartreuse4', foreground='black')
root.configure(bg='dark green')

# Configure a consistent font
default_font = ('Helvetica', 12)
style.configure('.', font=default_font)

# Title
title_label = ttk.Label(master=root, text='COOT Sorter', font=('Helvetica', 16))
title_label.pack(pady=(20, 10))

# Instructions
instructions_frame = ttk.Frame(master=root)
instructions_frame.pack(fill='x', padx=20, pady=10)
instructions_text = (
    "Select an Excel file (.xlsx) to sort students into trips. The file should consist of two sheets:\n\n"
    "Students: This sheet should include columns for Student ID, Gender, Team, POC, Dorm, Water, Tent, "
    "and preferences (A-J) which correspond to trip categories. Each row represents a student's information "
    "and preferences.\n\n"
    "Trips: This sheet should include columns for Trip, Category, Capacity, Water, and Tent. Each row "
    "represents a trip and its attributes.\n\n"
    "To process the file:\n"
    "1. Click 'Select File' to choose the Excel file.\n"
    "2. After selecting the file, its path will be displayed in the GUI.\n"
    "3. Click 'Process File' to assign students to trips based on their preferences and the trip capacities.\n"
    "4. Once processing is complete, you will be prompted to save the output. The result will be an Excel file with updated trip assignments and statistics."
)

instructions_label = ttk.Label(instructions_frame, text=instructions_text, wraplength=900)
instructions_label.pack(side='top', fill='x')

# Input
input_frame = ttk.Frame(master=root)
input_frame.pack(fill='x', padx=20, pady=10)

file_select = ttk.Button(input_frame, text="Select File", command=select_file)
file_select.pack(side='top', pady=(0, 10))

file_path_label = tk.Label(input_frame, text="", wraplength=700, justify='center', bg='chartreuse4')
file_path_label.pack(side='top', fill='x', anchor='center')

# Button to process the file
process_button = ttk.Button(input_frame, text="Process File", command=process_file)
process_button.pack(side='top', pady=(10, 20))

root.mainloop()
