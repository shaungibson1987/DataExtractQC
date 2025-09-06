import pandas as pd
import os
import sys
import traceback
from tqdm import tqdm

def log_error(message, error_log_path):
    with open(error_log_path, 'a', encoding='utf-8') as f:
        f.write(message + '\n')
        f.write(traceback.format_exc() + '\n')


def get_open_ends(df):
    cols = list(df.columns)
    # Try TESTJUMP, then TESTLANG, then outro
    if 'TESTJUMP' in cols:
        start_idx = cols.index('TESTJUMP') + 1
    elif 'TESTLANG' in cols:
        start_idx = cols.index('TESTLANG') + 1
    elif 'outro' in cols:
        start_idx = cols.index('outro') + 1
    else:
        print("Neither 'TESTJUMP', 'TESTLANG', nor 'outro' found as reference column.")
        return []
    open_end_cols = []
    for col in cols[start_idx:]:
        col_lower = col.lower()
        values = df[col].dropna().astype(str)
        values = [v for v in values if v.strip() != '']
        if not values:
            continue
        if '.oth' in col_lower or '._oth' in col_lower:
            open_end_cols.append(col)
            continue
        if any(not (v.startswith('_') or v[0].isdigit()) for v in values):
            open_end_cols.append(col)
    return open_end_cols

# Step 1: Ask for the input file (any path)
input_file = input('Enter the FULL path to your Excel (.xlsx) file (e.g., C:/Users/sgibson/Documents/mydata.xlsx): ').strip()
error_log_path = os.path.join(os.path.dirname(input_file), 'error_log.txt')

# Step 2: Read the file (.xlsx only)
try:
    print('Loading Excel file. This may take a while for large files...')
    df = pd.read_excel(input_file, dtype=str)
    print('Excel file loaded successfully.')
except Exception as e:
    log_error(f'Error reading Excel file: {e}', error_log_path)
    print(f'Error reading Excel file. See error_log.txt for details.')
    sys.exit(1)

# Step 2b: Ask if user wants to check for open ends
try:
    print('Checking for open ends is optional and may take a few moments if enabled.')
    check_open_ends = input('Do you want to check for open ends to add to the include file? (Y/N): ').strip().upper()
except Exception as e:
    log_error(f'Error during open ends prompt: {e}', error_log_path)
    print(f'Error during open ends prompt. See error_log.txt for details.')
    sys.exit(1)

# Step 3: Ask for the path to include.txt
try:
    print('Reading your include.txt file...')
    include_file = input('Enter the FULL path to your include.txt file (e.g., C:/Users/sgibson/DataExtractQC/P026776/include.txt): ').strip()
except Exception as e:
    log_error(f'Error during include.txt prompt: {e}', error_log_path)
    print(f'Error during include.txt prompt. See error_log.txt for details.')
    sys.exit(1)

# Step 3b: If yes, run open ends function and update include.txt
try:
    if check_open_ends == 'Y':
        print('Scanning for open ends. This may take a few moments...')
        open_end_cols = get_open_ends(df)
        try:
            with open(include_file, 'r', encoding='utf-8') as f:
                line = f.readline().strip()
                selected_columns = [col.strip() for col in line.split(';') if col.strip()]
        except Exception as e:
            log_error(f'Error reading include.txt: {e}', error_log_path)
            print(f'Error reading include.txt. See error_log.txt for details.')
            sys.exit(1)
        # Add new open ends, avoiding duplicates
        selected_columns += open_end_cols
        # Deduplicate while preserving order
        seen = set()
        deduped_columns = []
        for col in selected_columns:
            if col not in seen:
                deduped_columns.append(col)
                seen.add(col)
        # Update include.txt
        with open(include_file, 'w', encoding='utf-8') as f:
            f.write(';'.join(deduped_columns) + ';\n')
        print(f'Updated include.txt with open ends: {open_end_cols}')
        selected_columns = deduped_columns
    else:
        try:
            with open(include_file, 'r', encoding='utf-8') as f:
                line = f.readline().strip()
                selected_columns = [col.strip() for col in line.split(';') if col.strip()]
        except Exception as e:
            log_error(f'Error reading include.txt: {e}', error_log_path)
            print(f'Error reading include.txt. See error_log.txt for details.')
            sys.exit(1)
except Exception as e:
    log_error(f'Unexpected error handling include.txt: {e}', error_log_path)
    print(f'Unexpected error handling include.txt. See error_log.txt for details.')
    sys.exit(1)

try:
    print('Filtering columns and preparing for output...')
    all_columns = list(df.columns)
    selected_columns = [col for col in selected_columns if col in all_columns]
    if not selected_columns:
        print('No valid columns selected from include.txt. Exiting.')
        sys.exit(1)
    print(f'\nColumns included from include.txt: {selected_columns}')
except Exception as e:
    log_error(f'Error processing selected columns: {e}', error_log_path)
    print(f'Error processing selected columns. See error_log.txt for details.')
    sys.exit(1)

# Step 4: Ask for output directory
try:
    print('Please specify where to save your output files...')
    output_dir = input('Enter the FULL path to the folder where you want to save the output files (e.g., C:/Users/sgibson/Documents/Output): ').strip()
    if not os.path.isdir(output_dir):
        print(f'Output directory does not exist: {output_dir}')
        sys.exit(1)
except Exception as e:
    log_error(f'Error during output directory prompt: {e}', error_log_path)
    print(f'Error during output directory prompt. See error_log.txt for details.')
    sys.exit(1)

# Step 5: Find unique InterviewLanguage values
try:
    print('Scanning for unique InterviewLanguage values...')
    if 'InterviewLanguage' not in df.columns:
        print('InterviewLanguage column not found. Exiting.')
        sys.exit(1)
    languages = df['InterviewLanguage'].dropna().unique()
    print(f'\nLanguages found: {languages}')
except Exception as e:
    log_error(f'Error finding InterviewLanguage values: {e}', error_log_path)
    print(f'Error finding InterviewLanguage values. See error_log.txt for details.')
    sys.exit(1)

# Step 6: Create output files
try:
    print('Creating output files. This may take a few moments...')
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    logfile_path = os.path.join(output_dir, f'{base_name}_log.txt')
    log_lines = []
    log_lines.append(f'Columns kept: {selected_columns}')
    log_lines.append('Country\tRows\tCharacters')
    # Overall file (.xlsx)
    overall_file = os.path.join(output_dir, f'{base_name}__Overall.xlsx')
    df[selected_columns].to_excel(overall_file, index=False)
    print(f'Overall file created: {overall_file}')
    # Country/language files (.xlsx) with progress bar
    for lang in tqdm(languages, desc='Processing languages'):
        lang_df = df[df['InterviewLanguage'] == lang][selected_columns]
        out_file = os.path.join(output_dir, f'{base_name}__{lang}.xlsx')
        lang_df.to_excel(out_file, index=False)
        print(f'File created for {lang}: {out_file}')
        lang_chars = lang_df.astype(str).applymap(len).sum().sum()
        lang_rows = len(lang_df)
        log_lines.append(f'{lang}\t{lang_rows}\t{lang_chars}')
    # Save overall logfile
    with open(logfile_path, 'w', encoding='utf-8') as logf:
        for line in log_lines:
            logf.write(line + '\n')
    print(f'Logfile saved: {logfile_path}')
    print('\nDone!')
    print('\nSummary of files created:')
    print(f'- Overall file: {overall_file}')
    for lang in languages:
        out_file = os.path.join(output_dir, f'{base_name}__{lang}.xlsx')
        print(f'- Language file: {out_file}')
    print(f'- Log file: {logfile_path}')
    input('\nPress Enter to exit...')
except Exception as e:
    log_error(f'Error during output file creation: {e}', error_log_path)
    print(f'Error during output file creation. See error_log.txt for details.')
    sys.exit(1)
