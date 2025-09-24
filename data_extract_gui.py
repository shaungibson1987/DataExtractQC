import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from core import run_data_extract
from themes_list import THEMES
from status_handler import StatusHandler
from constants import LABEL_BROWSE, LABEL_MENU, LABEL_RUN_EXTRACTION, LABEL_DATA_FILE, LABEL_INCLUDE_FILE, LABEL_OUTPUT_FOLDER, LABEL_OPEN_ENDS, LABEL_AI_BOT_SEARCH, ERROR_SELECT_FILES, ERROR_GENERIC, STATUS_MESSAGES

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
	app.geometry("650x680")  # Increased height further for sentiment analysis controls

	# Variables for form fields and status
	excel_var = tk.StringVar()
	include_var = tk.StringVar()
	output_var = tk.StringVar()
	open_ends_var = tk.BooleanVar(value=False)
	ai_bot_search_var = tk.BooleanVar(value=False)
	status_var = tk.StringVar()
	word_file_var = tk.StringVar()
	duplicate_postcode_yob_var = tk.BooleanVar(value=False)
	length_check_var = tk.BooleanVar(value=False)
	length_multiplier_var = tk.StringVar(value="10")
	sentiment_analysis_var = tk.BooleanVar(value=False)
	sentiment_pos_threshold_var = tk.StringVar(value="0.6")
	sentiment_neg_threshold_var = tk.StringVar(value="0")

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
	lbl_file = tb.Label(form_frame, text=LABEL_DATA_FILE, anchor="w")
	lbl_file.pack(anchor="w")
	file_row = tb.Frame(form_frame)
	file_row.pack(fill="x", pady=(0, 10))
	excel_entry = tb.Entry(file_row, textvariable=excel_var)
	excel_entry.pack(side="left", fill="x", expand=True)
	btn_browse_excel = tb.Button(file_row, text=LABEL_BROWSE, bootstyle=PRIMARY, command=select_excel)
	btn_browse_excel.pack(side="left", padx=(8, 0))
	ToolTip(excel_entry, "Select the Excel data file (.xlsx) to process.")

	# Include.txt File
	lbl_include = tb.Label(form_frame, text=LABEL_INCLUDE_FILE, anchor="w")
	lbl_include.pack(anchor="w")
	include_row = tb.Frame(form_frame)
	include_row.pack(fill="x", pady=(0, 10))
	include_entry = tb.Entry(include_row, textvariable=include_var)
	include_entry.pack(side="left", fill="x", expand=True)
	btn_browse_include = tb.Button(include_row, text=LABEL_BROWSE, bootstyle=PRIMARY, command=select_include)
	btn_browse_include.pack(side="left", padx=(8, 0))
	ToolTip(include_entry, "Select the include.txt file listing columns to extract.")

	# Output Folder
	lbl_output = tb.Label(form_frame, text=LABEL_OUTPUT_FOLDER, anchor="w")
	lbl_output.pack(anchor="w")
	output_row = tb.Frame(form_frame)
	output_row.pack(fill="x", pady=(0, 10))
	output_entry = tb.Entry(output_row, textvariable=output_var)
	output_entry.pack(side="left", fill="x", expand=True)
	btn_browse_output = tb.Button(output_row, text=LABEL_BROWSE, bootstyle=PRIMARY, command=select_output)
	btn_browse_output.pack(side="left", padx=(8, 0))
	ToolTip(output_entry, "Choose the folder where output files will be saved.")


	# Options Section
	options_frame = tb.Frame(content_frame)
	options_frame.pack(pady=(0, 10), padx=20, fill="x")

	cb_open_ends = tb.Checkbutton(
		options_frame,
		text=LABEL_OPEN_ENDS,
		variable=open_ends_var,
		bootstyle="success-round-toggle"
	)
	cb_open_ends.pack(anchor="w", pady=(0, 16))  # Increased bottom padding

	# Length check toggle and multiplier
	length_check_row = tb.Frame(options_frame)
	cb_length_check = tb.Checkbutton(
		length_check_row,
		text="Enable length checks",
		variable=length_check_var,
		bootstyle="info-round-toggle"
	)
	cb_length_check.pack(side="left", anchor="w")

	tb.Label(length_check_row, text="Multiplier:").pack(side="left", padx=(10, 2))
	multiplier_entry = tb.Entry(length_check_row, textvariable=length_multiplier_var, width=5)
	multiplier_entry.pack(side="left")

	def toggle_multiplier_entry(*args):
		if length_check_var.get():
			multiplier_entry.config(state="normal")
		else:
			multiplier_entry.config(state="disabled")
	length_check_var.trace_add('write', toggle_multiplier_entry)
	toggle_multiplier_entry()

	length_check_row.pack(anchor="w", pady=(0, 16))

	# Duplicate postcode/yob toggle
	cb_duplicate_postcode_yob = tb.Checkbutton(
		options_frame,
		text="Check for duplicate postcode/YOB pairs",
		variable=duplicate_postcode_yob_var,
		bootstyle="warning-round-toggle"
	)
	cb_duplicate_postcode_yob.pack(anchor="w", pady=(0, 16))

	# Sentiment analysis toggle and thresholds
	sentiment_row = tb.Frame(options_frame)
	cb_sentiment = tb.Checkbutton(
		sentiment_row,
		text="Sentiment analysis on outro column",
		variable=sentiment_analysis_var,
		bootstyle="info-round-toggle"
	)
	cb_sentiment.pack(side="left", anchor="w")

	tb.Label(sentiment_row, text="Pos:").pack(side="left", padx=(10, 2))
	pos_threshold_entry = tb.Entry(sentiment_row, textvariable=sentiment_pos_threshold_var, width=5)
	pos_threshold_entry.pack(side="left")

	tb.Label(sentiment_row, text="Neg:").pack(side="left", padx=(5, 2))
	neg_threshold_entry = tb.Entry(sentiment_row, textvariable=sentiment_neg_threshold_var, width=5)
	neg_threshold_entry.pack(side="left")

	def toggle_sentiment_entries(*args):
		if sentiment_analysis_var.get():
			pos_threshold_entry.config(state="normal")
			neg_threshold_entry.config(state="normal")
		else:
			pos_threshold_entry.config(state="disabled")
			neg_threshold_entry.config(state="disabled")
	sentiment_analysis_var.trace_add('write', toggle_sentiment_entries)
	toggle_sentiment_entries()

	sentiment_row.pack(anchor="w", pady=(0, 16))

	# Frame for word search toggle and word file input
	word_search_frame = tb.Frame(options_frame)
	word_search_frame.pack(anchor="w", fill="x")

	cb_ai_bot_search = tb.Checkbutton(
		word_search_frame,
		text=LABEL_AI_BOT_SEARCH,
		variable=ai_bot_search_var,
		bootstyle="success-round-toggle"
	)
	cb_ai_bot_search.pack(anchor="w", pady=(0, 8))  # Add bottom padding for word file input

	# Frame for word file input (initially hidden)
	word_file_row = tb.Frame(word_search_frame)
	word_file_entry = tb.Entry(word_file_row, textvariable=word_file_var, width=30)
	btn_browse_word_file = tb.Button(word_file_row, text=LABEL_BROWSE, bootstyle=PRIMARY, command=lambda: select_word_file())
	word_file_entry.pack(side="left", fill="x", expand=True)
	btn_browse_word_file.pack(side="left", padx=(8, 0))
	ToolTip(word_file_entry, "Select a .txt file with search words (one per line).")

	def select_word_file():
		file_path = filedialog.askopenfilename(
			filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
			title="Select word search .txt file"
		)
		if file_path:
			word_file_var.set(file_path)

	def toggle_word_file_row(*args):
		if ai_bot_search_var.get():
			word_file_row.pack(anchor="w", pady=(4, 0), fill="x")
		else:
			word_file_row.pack_forget()

	ai_bot_search_var.trace_add('write', toggle_word_file_row)
	# Set initial visibility
	toggle_word_file_row()

	def on_run():
		input_file = excel_var.get()
		include_file = include_var.get()
		output_dir = output_var.get()
		check_open_ends = open_ends_var.get()
		check_ai_bot_search = ai_bot_search_var.get()
		word_file = word_file_var.get()
		if not input_file or not include_file or not output_dir:
			messagebox.showerror("Error", ERROR_SELECT_FILES)
			return
		if check_ai_bot_search and not word_file:
			messagebox.showerror("Error", "Please select a .txt file with search words.")
			return

		check_duplicate_postcode_yob = duplicate_postcode_yob_var.get()
		check_length = length_check_var.get()
		try:
			length_multiplier = int(length_multiplier_var.get())
		except Exception:
			length_multiplier = 10

		# Get sentiment analysis parameters
		check_sentiment = sentiment_analysis_var.get()
		try:
			sentiment_pos_threshold = float(sentiment_pos_threshold_var.get())
		except Exception:
			sentiment_pos_threshold = 0.05
		try:
			sentiment_neg_threshold = float(sentiment_neg_threshold_var.get())
		except Exception:
			sentiment_neg_threshold = -0.05

		# Step-by-step status with color and progress
		def status_callback(msg):
			status.set_status(msg, "blue")

		status.set_status(STATUS_MESSAGES['load_excel'], "blue")
		# Pass word_file, duplicate toggle, length check toggle, multiplier, and sentiment parameters to run_data_extract
		success = run_data_extract(
			input_file, include_file, output_dir, check_open_ends, check_ai_bot_search, word_file, status_callback=status_callback,
			check_duplicate_postcode_yob=check_duplicate_postcode_yob, check_length=check_length, length_multiplier=length_multiplier,
			check_sentiment=check_sentiment, sentiment_pos_threshold=sentiment_pos_threshold, sentiment_neg_threshold=sentiment_neg_threshold
		)
		if success:
			status.set_status("Done!!! Your files are ready, check your output folder.", "green")
			messagebox.showinfo("Done", "Processing complete! See output folder for results.")
		else:
			status.set_status(ERROR_GENERIC, "red")
			messagebox.showerror("Error", ERROR_GENERIC)

	# Run Button
	btn_run = tb.Button(content_frame, text=LABEL_RUN_EXTRACTION, bootstyle=SUCCESS, command=on_run)
	btn_run.pack(pady=20)

	# Status bar at the very bottom (fixed using pack)
	status_bar = tb.Frame(app, bootstyle="light")
	status_bar.pack(side="bottom", fill="x")
	status_label = tb.Label(status_bar, textvariable=status_var, bootstyle="secondary", anchor="w", font=("Segoe UI", 12))  # Larger font
	status_label.pack(fill="x", padx=10, pady=2)

	status = StatusHandler(status_var, status_label, app)

	# Set grid column weights: label and button columns do not expand, entry column expands
	content_frame.columnconfigure(0, weight=0, minsize=0)
	content_frame.columnconfigure(1, weight=1)
	content_frame.columnconfigure(2, weight=0)

	app.mainloop()

if __name__ == "__main__":
	main()
