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
    if 'TESTJUMP' in cols:
        start_idx = cols.index('TESTJUMP') + 1
    elif 'TESTLANG' in cols:
        start_idx = cols.index('TESTLANG') + 1
    elif 'outro' in cols:
        start_idx = cols.index('outro') + 1
    else:
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

def run_data_extract(input_file, include_file, output_dir, check_open_ends=True, error_log_path=None, status_callback=None):
    import time
    from datetime import datetime

    if error_log_path is None:
        error_log_path = os.path.join(os.path.dirname(input_file), 'error_log.txt')

    start_time = datetime.now()
    start_ts = time.time()

    # Step 1: Load Excel file
    try:
        if status_callback:
            status_callback('Loading Excel file...')
        df = pd.read_excel(input_file, dtype=str)
    except Exception as e:
        log_error(f'Error reading Excel file: {e}', error_log_path)
        if status_callback:
            status_callback('Error reading Excel file.')
        return False

    # Step 2: Handle include.txt and open ends
    try:
        open_end_cols = []
        original_include_columns = []
        include_with_opens_path = os.path.join(output_dir, 'Include_withOpens.txt')
        if check_open_ends:
            if status_callback:
                status_callback('Scanning for open ends...')
            open_end_cols = get_open_ends(df)
            try:
                with open(include_file, 'r', encoding='utf-8') as f:
                    line = f.readline().strip()
                    original_include_columns = [col.strip() for col in line.split(';') if col.strip()]
            except Exception as e:
                log_error(f'Error reading include.txt: {e}', error_log_path)
                if status_callback:
                    status_callback('Error reading include.txt.')
                return False
            selected_columns = original_include_columns + open_end_cols
            # Deduplicate while preserving order
            seen = set()
            deduped_columns = []
            for col in selected_columns:
                if col not in seen:
                    deduped_columns.append(col)
                    seen.add(col)
            # Save the new include file with open ends in the output directory
            try:
                with open(include_with_opens_path, 'w', encoding='utf-8') as f:
                    f.write(';'.join(deduped_columns) + ';\n')
            except Exception as e:
                log_error(f'Error writing Include_withOpens.txt: {e}', error_log_path)
                if status_callback:
                    status_callback('Error writing Include_withOpens.txt.')
                return False
            selected_columns = deduped_columns
        else:
            try:
                with open(include_file, 'r', encoding='utf-8') as f:
                    line = f.readline().strip()
                    original_include_columns = [col.strip() for col in line.split(';') if col.strip()]
            except Exception as e:
                log_error(f'Error reading include.txt: {e}', error_log_path)
                if status_callback:
                    status_callback('Error reading include.txt.')
                return False
            selected_columns = original_include_columns
    except Exception as e:
        log_error(f'Unexpected error handling include.txt: {e}', error_log_path)
        if status_callback:
            status_callback('Unexpected error handling include.txt.')
        return False

    # Step 3: Filter columns
    try:
        if status_callback:
            status_callback('Filtering columns and preparing for output...')
        all_columns = list(df.columns)
        selected_columns = [col for col in selected_columns if col in all_columns]
        if not selected_columns:
            if status_callback:
                status_callback('No valid columns selected from include.txt.')
            return False
    except Exception as e:
        log_error(f'Error processing selected columns: {e}', error_log_path)
        if status_callback:
            status_callback('Error processing selected columns.')
        return False

    # Step 4: Check output directory
    try:
        if not os.path.isdir(output_dir):
            if status_callback:
                status_callback(f'Output directory does not exist: {output_dir}')
            return False
    except Exception as e:
        log_error(f'Error during output directory prompt: {e}', error_log_path)
        if status_callback:
            status_callback('Error during output directory prompt.')
        return False

    # Step 5: Find unique InterviewLanguage values
    try:
        if status_callback:
            status_callback('Scanning for unique InterviewLanguage values...')
        if 'InterviewLanguage' not in df.columns:
            if status_callback:
                status_callback('InterviewLanguage column not found.')
            return False
        languages = df['InterviewLanguage'].dropna().unique()
    except Exception as e:
        log_error(f'Error finding InterviewLanguage values: {e}', error_log_path)
        if status_callback:
            status_callback('Error finding InterviewLanguage values.')
        return False

    # Step 6: Create output files and enhanced log
    try:
        if status_callback:
            status_callback('Creating output files...')
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        logfile_path = os.path.join(output_dir, f'{base_name}_log.txt')
        output_files = []
        char_counts = {}
        outro_char_counts = {}
        # Overall file
        overall_file = os.path.join(output_dir, f'{base_name}__Overall.xlsx')
        df[selected_columns].to_excel(overall_file, index=False)
        output_files.append(overall_file)
        char_counts[overall_file] = df[selected_columns].astype(str).applymap(len).sum().sum()
        outro_char_counts[overall_file] = df['outro'].astype(str).apply(len).sum() if 'outro' in df.columns else 0
        # Per-language files
        for lang in languages:
            lang_df = df[df['InterviewLanguage'] == lang][selected_columns]
            out_file = os.path.join(output_dir, f'{base_name}__{lang}.xlsx')
            lang_df.to_excel(out_file, index=False)
            output_files.append(out_file)
            char_counts[out_file] = lang_df.astype(str).applymap(len).sum().sum()
            outro_char_counts[out_file] = lang_df['outro'].astype(str).apply(len).sum() if 'outro' in lang_df.columns else 0
        end_time = datetime.now()
        end_ts = time.time()
        runtime = end_ts - start_ts
        # Open ends not in original include file
        new_open_ends = [col for col in open_end_cols if col not in original_include_columns]
        # Compose log
        log_lines = []
        log_lines.append(f"Start time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        log_lines.append(f"End time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        log_lines.append(f"Total running time in seconds: {runtime:.2f}")
        log_lines.append("")
        log_lines.append(f"Number of output files created: {len(output_files)}")
        log_lines.append(f"Languages in the data file: {', '.join(str(l) for l in languages)}")
        log_lines.append("")
        log_lines.append("Output files:")
        for f in output_files:
            log_lines.append(f"  {f}")
        log_lines.append("")
        log_lines.append(f"Number of open ends detected that weren't originally in the include file: {len(new_open_ends)}")
        if new_open_ends:
            log_lines.append("Names of new open end columns:")
            for col in new_open_ends:
                log_lines.append(f"  {col}")
        log_lines.append("")
        log_lines.append("Character counts per file:")
        for f in output_files:
            log_lines.append(f"  {os.path.basename(f)}: {char_counts[f]}")
        log_lines.append("")
        log_lines.append("Number of characters in the 'outro' column of each file:")
        for f in output_files:
            log_lines.append(f"  {os.path.basename(f)}: {outro_char_counts[f]}")
        log_lines.append("")
        with open(logfile_path, 'w', encoding='utf-8') as logf:
            for line in log_lines:
                logf.write(line + '\n')
        if status_callback:
            status_callback('Done! Files created:')
            status_callback(f'- Overall file: {overall_file}')
            for lang in languages:
                out_file = os.path.join(output_dir, f'{base_name}__{lang}.xlsx')
                status_callback(f'- Language file: {out_file}')
            status_callback(f'- Log file: {logfile_path}')
        return True
    except Exception as e:
        log_error(f'Error during output file creation: {e}', error_log_path)
        if status_callback:
            status_callback('Error during output file creation.')
        return False