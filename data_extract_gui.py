
import tkinter as tk
from tkinter import filedialog, messagebox
from core import run_data_extract

def select_excel():
	path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
	excel_var.set(path)

def select_include():
	path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
	include_var.set(path)

def select_output():
	path = filedialog.askdirectory()
	output_var.set(path)


root = tk.Tk()
root.title("DataExtractQC GUI")
root.geometry("500x400")


excel_var = tk.StringVar()
include_var = tk.StringVar()
output_var = tk.StringVar()
open_ends_var = tk.BooleanVar()
status_var = tk.StringVar()


# Data file (.xlsx)

tk.Label(root, text="Data File (.xlsx):").pack()
tk.Entry(root, textvariable=excel_var, width=50).pack()
tk.Button(root, text="Browse", command=select_excel).pack()
tk.Label(root, text="").pack()  # Spacer

# Include.txt file

tk.Label(root, text="Include.txt file:").pack()
tk.Entry(root, textvariable=include_var, width=50).pack()
tk.Button(root, text="Browse", command=select_include).pack()
tk.Label(root, text="").pack()  # Spacer

# Checkbox for open ends (after include file)

tk.Checkbutton(root, text="Automatically search the data file for open ends.", variable=open_ends_var).pack(pady=5)
tk.Label(root, text="").pack()  # Spacer

# Output folder

tk.Label(root, text="Output folder:").pack()
tk.Entry(root, textvariable=output_var, width=50).pack()
tk.Button(root, text="Browse", command=select_output).pack()
tk.Label(root, text="").pack()  # Spacer


# Status label
tk.Label(root, textvariable=status_var, fg="blue").pack(pady=5)

# Run button
def on_run():
	input_file = excel_var.get()
	include_file = include_var.get()
	output_dir = output_var.get()
	check_open_ends = open_ends_var.get()
	if not input_file or not include_file or not output_dir:
		messagebox.showerror("Error", "Please select all required files and output folder.")
		return
	def status_callback(msg):
		status_var.set(msg)
		root.update_idletasks()
	success = run_data_extract(input_file, include_file, output_dir, check_open_ends, status_callback=status_callback)
	if success:
		messagebox.showinfo("Done", "Processing complete! See output folder for results.")
	else:
		messagebox.showerror("Error", "An error occurred. See error_log.txt for details.")

tk.Button(root, text="Run", command=on_run).pack(pady=10)


root.mainloop()
