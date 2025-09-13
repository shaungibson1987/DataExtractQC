import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from core import run_data_extract
from themes_list import THEMES
from status_handler import StatusHandler

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

def main():
	# Default theme
	current_theme = {"name": "yeti"}
	app = tb.Window(themename=current_theme["name"])
	app.title("Data Extract QC")
	app.geometry("650x500")

	# Variables for form fields and status
	excel_var = tk.StringVar()
	include_var = tk.StringVar()
	output_var = tk.StringVar()
	open_ends_var = tk.BooleanVar()
	status_var = tk.StringVar()

	# --- File/folder selection functions ---
	def select_excel():
		file_path = filedialog.askopenfilename(
			filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
			title="Select Excel Data File"
		)
		if file_path:
			excel_var.set(file_path)

	def select_include():
		file_path = filedialog.askopenfilename(
			filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
			title="Select include.txt File"
		)
		if file_path:
			include_var.set(file_path)

	def select_output():
		folder_path = filedialog.askdirectory(title="Select Output Folder")
		if folder_path:
			output_var.set(folder_path)

	# --- Classic menu bar for theme selection with padding and visual separation ---
	def set_theme(theme_name):
		try:
			app.style.theme_use(theme_name)
			current_theme["name"] = theme_name
		except Exception as e:
			messagebox.showerror("Theme Error", f"Could not set theme: {theme_name}\n{e}")

	# Create a content frame for all main widgets (grid layout)
	content_frame = tb.Frame(app)
	content_frame.pack(fill="both", expand=True)

	# # Add a top frame for padding and visual separation inside content_frame
	# top_frame = tb.Frame(content_frame, bootstyle="secondary", padding=(0, 0, 0, 2))
	# top_frame.grid(row=0, column=0, columnspan=3, sticky="ew")
	# Add a subtle border or background to the menu bar
	menubar = tk.Menu(app, bg="#f0f0f0", relief="flat", bd=1)
	# Themes submenu
	themes_submenu = tk.Menu(menubar, tearoff=0, bg="#f0f0f0")
	for theme in THEMES:
		themes_submenu.add_command(
			label=theme,
			command=lambda t=theme: set_theme(t)
		)
	# Menu -> Themes -> [theme list]
	menu_dropdown = tk.Menu(menubar, tearoff=0, bg="#f0f0f0")
	menu_dropdown.add_cascade(label="Themes", menu=themes_submenu)
	menubar.add_cascade(label="Menu", menu=menu_dropdown)
	app.config(menu=menubar)

	# --- Modern Vertical Layout ---
	# Title & Description
	title_label = tb.Label(content_frame, text="QC Data Extract", font=("Segoe UI", 18, "bold"), anchor="center", justify="center")
	title_label.pack(pady=(20, 5))

	desc_text = (
		"Easily extract, filter, and split survey data into organized Excel files.\n"
		"Allows creation of translation files for Google Translate."
	)
	desc_label = tb.Label(content_frame, text=desc_text, font=("Segoe UI", 10), anchor="center", justify="center")
	desc_label.pack(pady=(0, 20))

	# Form Frame
	form_frame = tb.Frame(content_frame)
	form_frame.pack(pady=10, padx=20, fill="x")

	# Data File
	lbl_file = tb.Label(form_frame, text="Data File (.xlsx):", anchor="w")
	lbl_file.pack(anchor="w")
	file_row = tb.Frame(form_frame)
	file_row.pack(fill="x", pady=(0, 10))
	excel_entry = tb.Entry(file_row, textvariable=excel_var)
	excel_entry.pack(side="left", fill="x", expand=True)
	btn_browse_excel = tb.Button(file_row, text="Browse", bootstyle=PRIMARY, command=select_excel)
	btn_browse_excel.pack(side="left", padx=(8, 0))
	ToolTip(excel_entry, "Select the Excel data file (.xlsx) to process.")

	# Include.txt File
	lbl_include = tb.Label(form_frame, text="Include.txt file:", anchor="w")
	lbl_include.pack(anchor="w")
	include_row = tb.Frame(form_frame)
	include_row.pack(fill="x", pady=(0, 10))
	include_entry = tb.Entry(include_row, textvariable=include_var)
	include_entry.pack(side="left", fill="x", expand=True)
	btn_browse_include = tb.Button(include_row, text="Browse", bootstyle=PRIMARY, command=select_include)
	btn_browse_include.pack(side="left", padx=(8, 0))
	ToolTip(include_entry, "Select the include.txt file listing columns to extract.")

	# Output Folder
	lbl_output = tb.Label(form_frame, text="Output folder:", anchor="w")
	lbl_output.pack(anchor="w")
	output_row = tb.Frame(form_frame)
	output_row.pack(fill="x", pady=(0, 10))
	output_entry = tb.Entry(output_row, textvariable=output_var)
	output_entry.pack(side="left", fill="x", expand=True)
	btn_browse_output = tb.Button(output_row, text="Browse", bootstyle=PRIMARY, command=select_output)
	btn_browse_output.pack(side="left", padx=(8, 0))
	ToolTip(output_entry, "Choose the folder where output files will be saved.")

	# Options Section
	options_frame = tb.Frame(content_frame)
	options_frame.pack(pady=(0, 10), padx=20, fill="x")
	cb_open_ends = tb.Checkbutton(
		options_frame,
		text="Automatically search the data file for open ends.",
		variable=open_ends_var,
		bootstyle="success-round-toggle"
	)
	cb_open_ends.pack(anchor="w")

	def on_run():
		input_file = excel_var.get()
		include_file = include_var.get()
		output_dir = output_var.get()
		check_open_ends = open_ends_var.get()
		# check_words = check_words_var.get()
		# words_file = words_file_var.get()
		if not input_file or not include_file or not output_dir:
			messagebox.showerror("Error", "Please select all required files and output folder.")
			return

		# Step-by-step status with color and progress
		def status_callback(msg):
			status.set_status(msg, "blue")

		status.set_status("Step 1 of 5: Loading your Excel file...", "blue")
		# Pass check_words and words_file to run_data_extract if needed
		success = run_data_extract(
			input_file, include_file, output_dir, check_open_ends, status_callback=status_callback
		)
		if success:
			status.set_status("Step 5 of 5: Done! Your files are ready.", "green")
			messagebox.showinfo("Done", "Processing complete! See output folder for results.")
		else:
			status.set_status("Error: An error occurred. See error_log.txt for details.", "red")
			messagebox.showerror("Error", "An error occurred. See error_log.txt for details.")

	# Run Button
	btn_run = tb.Button(content_frame, text="Run Extraction", bootstyle=SUCCESS, command=on_run)
	btn_run.pack(pady=20)

	# Status bar at the very bottom (fixed using pack)
	status_bar = tb.Frame(app, bootstyle="light")
	status_bar.pack(side="bottom", fill="x")
	status_label = tb.Label(status_bar, textvariable=status_var, bootstyle="secondary", anchor="w")
	status_label.pack(fill="x", padx=10, pady=2)

	status = StatusHandler(status_var, status_label, app)

	# Set grid column weights: label and button columns do not expand, entry column expands
	content_frame.columnconfigure(0, weight=0, minsize=0)
	content_frame.columnconfigure(1, weight=1)
	content_frame.columnconfigure(2, weight=0)

	app.mainloop()

if __name__ == "__main__":
	main()
