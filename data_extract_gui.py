# Tooltip helper class
class ToolTip(object):
	def __init__(self, widget, text):
		self.widget = widget
		self.text = text
		self.tipwindow = None
		widget.bind("<Enter>", self.show_tip)
		widget.bind("<Leave>", self.hide_tip)

	def show_tip(self, event=None):
		if self.tipwindow or not self.text:
			return
		x, y, _, cy = self.widget.bbox("insert") if self.widget.winfo_class() == 'Entry' else (0, 0, 0, 0)
		x = x + self.widget.winfo_rootx() + 25
		y = y + cy + self.widget.winfo_rooty() + 20
		self.tipwindow = tw = tk.Toplevel(self.widget)
		tw.wm_overrideredirect(True)
		tw.wm_geometry(f"+{x}+{y}")
		label = tk.Label(tw, text=self.text, justify='left',
						 background="#ffffe0", relief='solid', borderwidth=1,
						 font=("tahoma", "8", "normal"))
		label.pack(ipadx=1)

	def hide_tip(self, event=None):
		tw = self.tipwindow
		self.tipwindow = None
		if tw:
			tw.destroy()

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
excel_entry = tk.Entry(root, textvariable=excel_var, width=50)
excel_entry.pack()
ToolTip(excel_entry, "Select the Excel data file (.xlsx) to process.")
tk.Button(root, text="Browse", command=select_excel).pack()
tk.Label(root, text="").pack()  # Spacer

# Include.txt file
tk.Label(root, text="Include.txt file:").pack()
include_entry = tk.Entry(root, textvariable=include_var, width=50)
include_entry.pack()
ToolTip(include_entry, "Select the include.txt file listing columns to extract.")
tk.Button(root, text="Browse", command=select_include).pack()
tk.Label(root, text="").pack()  # Spacer

# Checkbox for open ends (after include file)

tk.Checkbutton(root, text="Automatically search the data file for open ends.", variable=open_ends_var).pack(pady=5)
tk.Label(root, text="").pack()  # Spacer

# Output folder
tk.Label(root, text="Output folder:").pack()
output_entry = tk.Entry(root, textvariable=output_var, width=50)
output_entry.pack()
ToolTip(output_entry, "Choose the folder where output files will be saved.")
tk.Button(root, text="Browse", command=select_output).pack()
tk.Label(root, text="").pack()  # Spacer


# Status label (keep a reference for color changes)
status_label = tk.Label(root, textvariable=status_var, fg="blue")
status_label.pack(pady=5)

# Run button

def set_status(msg, color="blue"):
	status_var.set(msg)
	status_label.config(fg=color)
	root.update_idletasks()

def on_run():
	input_file = excel_var.get()
	include_file = include_var.get()
	output_dir = output_var.get()
	check_open_ends = open_ends_var.get()
	if not input_file or not include_file or not output_dir:
		messagebox.showerror("Error", "Please select all required files and output folder.")
		return

	# Step-by-step status with color and progress
	def status_callback(msg):
		set_status(msg, "blue")

	set_status("Step 1 of 5: Loading your Excel file...", "blue")
	success = run_data_extract(
		input_file, include_file, output_dir, check_open_ends, status_callback=status_callback
	)
	if success:
		set_status("Step 5 of 5: Done! Your files are ready.", "green")
		messagebox.showinfo("Done", "Processing complete! See output folder for results.")
	else:
		set_status("Error: An error occurred. See error_log.txt for details.", "red")
		messagebox.showerror("Error", "An error occurred. See error_log.txt for details.")

tk.Button(root, text="Run", command=on_run).pack(pady=10)


root.mainloop()
