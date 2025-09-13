# constants.py

# Default file names - used in multiple modules
ERROR_LOG_FILENAME = 'error_log.txt'
INCLUDE_WITH_OPENS_FILENAME = 'Include_withOpens.txt'
LOG_SUFFIX = '_log.txt'
OVERALL_SUFFIX = '__Overall.xlsx'
LANGUAGE_SUFFIX_TEMPLATE = '__{lang}.xlsx'

# Ignore set for open ends - used in GetOpenEnds.py
IGNORE_SET = {"yes", "no", "dontknow", "_ref"}

# Status messages (optional, can be expanded) - used in core.py and data_extract_gui.py
STATUS_MESSAGES = {
    'load_excel': 'Step 1 of 6: Loading your Excel file... This may take a minute for large data files.',
    'scan_open_ends': 'Step 2 of 6: Scanning for open-ended questions in your data...',
    'filter_columns': 'Step 3 of 6: Filtering columns and preparing output...',
    'scan_languages': 'Step 4 of 6: Scanning for unique InterviewLanguage values...',
    'create_files': 'Step 5 of 6: Creating output files...',
    'word_search': 'Step 6 of 6: Running word search and highlighting matches...'
}

# UI Labels - used in data_extract_gui.py
LABEL_BROWSE = "Browse"
LABEL_MENU = "Menu"
LABEL_RUN_EXTRACTION = "Run Extraction"
LABEL_DATA_FILE = "Data File (.xlsx):"
LABEL_INCLUDE_FILE = "Include.txt file:"
LABEL_OUTPUT_FOLDER = "Output folder:"
LABEL_OPEN_ENDS = "Automatically search the data file for open ends."
LABEL_AI_BOT_SEARCH = "Search file for specific words (words.txt)"

# Error Messages - used in data_extract_gui.py and core.py
ERROR_SELECT_FILES = "Please select all required files and output folder."
ERROR_GENERIC = "An error occurred. See error_log.txt for details."