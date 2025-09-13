# constants.py

# Default file names
ERROR_LOG_FILENAME = 'error_log.txt'
INCLUDE_WITH_OPENS_FILENAME = 'Include_withOpens.txt'
LOG_SUFFIX = '_log.txt'
OVERALL_SUFFIX = '__Overall.xlsx'
LANGUAGE_SUFFIX_TEMPLATE = '__{lang}.xlsx'

# Ignore set for open ends
IGNORE_SET = {"yes", "no", "dontknow", "_ref"}

# Status messages (optional, can be expanded)
STATUS_MESSAGES = {
    'load_excel': 'Step 1 of 5: Loading your Excel file...',
    'scan_open_ends': 'Step 2 of 5: Scanning for open-ended questions in your data...',
    'filter_columns': 'Step 3 of 5: Filtering columns and preparing output...',
    'scan_languages': 'Step 4 of 5: Scanning for unique InterviewLanguage values...',
    'create_files': 'Step 5 of 5: Creating output files...'
}

# UI Labels
LABEL_BROWSE = "Browse"
LABEL_MENU = "Menu"
LABEL_RUN_EXTRACTION = "Run Extraction"
LABEL_DATA_FILE = "Data File (.xlsx):"
LABEL_INCLUDE_FILE = "Include.txt file:"
LABEL_OUTPUT_FOLDER = "Output folder:"
LABEL_OPEN_ENDS = "Automatically search the data file for open ends."

# Error Messages
ERROR_SELECT_FILES = "Please select all required files and output folder."
ERROR_GENERIC = "An error occurred. See error_log.txt for details."
