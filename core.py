import pandas as pd
import os
from pathlib import Path
import sys
import traceback
from tqdm import tqdm
from GetOpenEnds import get_open_ends
from error_logging import log_error
from constants import ERROR_LOG_FILENAME, INCLUDE_WITH_OPENS_FILENAME, LOG_SUFFIX, OVERALL_SUFFIX, LANGUAGE_SUFFIX_TEMPLATE, IGNORE_SET, STATUS_MESSAGES
from debug_logging import debug

def run_data_extract(input_file, include_file, output_dir, check_open_ends=True, check_ai_bot_search=False, error_log_path=None, status_callback=None):
    import time
    from datetime import datetime

    if error_log_path is None:
        error_log_path = os.path.join(os.path.dirname(input_file), ERROR_LOG_FILENAME)

    start_time = datetime.now()
    start_ts = time.time()

    # Step 1: Load Excel file
    try:
        if status_callback:
            status_callback(STATUS_MESSAGES['load_excel'])
        df = pd.read_excel(input_file, dtype=str)
    except Exception as e:
        log_error(f'Error reading Excel file: {e}', error_log_path)
        if status_callback:
            status_callback('Error: Could not read the Excel file. Please check the file and try again.')
        return False

    # Step 2: Handle include.txt and open ends
    try:
        open_end_cols = []
        original_include_columns = []
        output_dir = Path(output_dir)
        include_with_opens_path = output_dir / INCLUDE_WITH_OPENS_FILENAME
        if check_open_ends:
            if status_callback:
                status_callback(STATUS_MESSAGES['scan_open_ends'])
            open_end_cols = get_open_ends(df)
            try:
                with open(include_file, 'r', encoding='utf-8') as f:
                    original_include_columns = [line.strip() for line in f if line.strip()]
            except Exception as e:
                log_error(f'Error reading include.txt: {e}', error_log_path)
                if status_callback:
                    status_callback('Error: Could not read the include.txt file. Please check the file and try again.')
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
                    for col in deduped_columns:
                        f.write(col + '\n')
            except Exception as e:
                log_error(f'Error writing Include_withOpens.txt: {e}', error_log_path)
                if status_callback:
                    status_callback('Error: Could not write Include_withOpens.txt in the output folder.')
                return False
            selected_columns = deduped_columns
        else:
            try:
                with open(include_file, 'r', encoding='utf-8') as f:
                    original_include_columns = [line.strip() for line in f if line.strip()]
            except Exception as e:
                log_error(f'Error reading include.txt: {e}', error_log_path)
                if status_callback:
                    status_callback('Error: Could not read the include.txt file. Please check the file and try again.')
                return False
            selected_columns = original_include_columns
    except Exception as e:
        log_error(f'Unexpected error handling include.txt: {e}', error_log_path)
        if status_callback:
            status_callback('Error: Unexpected error handling include.txt.')
        return False

    # Step 3: Filter columns
    try:
        if status_callback:
            status_callback(STATUS_MESSAGES['filter_columns'])
        all_columns = list(df.columns)
        selected_columns = [col for col in selected_columns if col in all_columns]
        if not selected_columns:
            if status_callback:
                status_callback('No valid columns selected from include.txt.')
            return False
    except Exception as e:
        log_error(f'Error processing selected columns: {e}', error_log_path)
        if status_callback:
            status_callback('Error: Problem processing selected columns.')
        return False

    # Step 4: Check output directory
    try:
        if not os.path.isdir(output_dir):
            if status_callback:
                status_callback(f'Error: Output directory does not exist: {output_dir}')
            return False
    except Exception as e:
        log_error(f'Error during output directory prompt: {e}', error_log_path)
        if status_callback:
            status_callback('Error: Problem with output directory.')
        return False

    # Step 5: Find unique InterviewLanguage values
    try:
        if status_callback:
            status_callback(STATUS_MESSAGES['scan_languages'])
        if 'InterviewLanguage' not in df.columns:
            if status_callback:
                status_callback('InterviewLanguage column not found.')
            return False
        languages = df['InterviewLanguage'].dropna().unique()
        debug(f"Languages found: {languages}")
    except Exception as e:
        log_error(f'Error finding InterviewLanguage values: {e}', error_log_path)
        if status_callback:
            status_callback('Error: Could not find InterviewLanguage values.')
        return False

    # Step 6: Create output files and enhanced log
    try:
        if status_callback:
            status_callback(STATUS_MESSAGES['create_files'])
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        logfile_path = os.path.join(output_dir, f'{base_name}{LOG_SUFFIX}')
        output_files = []
        char_counts = {}
        outro_char_counts = {}
        # Overall file
        overall_file = os.path.join(output_dir, f'{base_name}{OVERALL_SUFFIX}')
        df[selected_columns].to_excel(overall_file, index=False)
        output_files.append(overall_file)
        char_counts[overall_file] = df[selected_columns].astype(str).apply(lambda col: col.map(len)).sum().sum()
        outro_char_counts[overall_file] = df['outro'].astype(str).apply(len).sum() if 'outro' in df.columns else 0

        # --- AI/BOT/CHATBOT word search post-processing ---
        word_search_serials = None
        if check_ai_bot_search:
            import re
            search_words = ["AI", "CHATBOT", "BOT"]
            debug(f"[AI/BOT Search] Search words: {search_words}")
            try:
                df_overall = pd.read_excel(overall_file, dtype=str)
                # Log first 5 outro values
                if 'outro' in df_overall.columns:
                    outro_vals = df_overall['outro'].dropna().astype(str).head(5).tolist()
                    debug(f"[AI/BOT Search] First 5 outro values: {outro_vals}")
                else:
                    debug("[AI/BOT Search] No 'outro' column found in overall file.")
                matches = []
                flagged_serials = set()
                search_words_lower = [w.lower() for w in search_words]
                for idx, row in df_overall.iterrows():
                    for col in df_overall.columns:
                        cell = str(row[col]) if not pd.isna(row[col]) else ""
                        # Tokenize using Unicode word characters
                        tokens = re.findall(r'\w+', cell, flags=re.UNICODE)
                        tokens_lower = [t.lower() for t in tokens]
                        for word in search_words_lower:
                            if word in tokens_lower:
                                respondent_serial = row.get('serial', idx)
                                matches.append((respondent_serial, col, cell))
                                flagged_serials.add(respondent_serial)
                                break  # Only log once per cell
                word_search_serials = sorted(flagged_serials, key=lambda x: str(x))
                if matches:
                    debug(f"[AI/BOT Search] Matches found:")
                    for respondent_serial, col, cell in matches:
                        debug(f"  Respondent: {respondent_serial}, Column: {col}, Value: {cell}")
                else:
                    debug("[AI/BOT Search] No matches found.")
            except Exception as e:
                debug(f"[AI/BOT Search] Error during search: {e}")
        # Per-language files
        for lang in languages:
            lang_df = df[df['InterviewLanguage'] == lang][selected_columns]
            out_file = os.path.join(output_dir, f'{base_name}{LANGUAGE_SUFFIX_TEMPLATE.format(lang=lang)}')
            lang_df.to_excel(out_file, index=False)
            output_files.append(out_file)
            char_counts[out_file] = lang_df.astype(str).apply(lambda col: col.map(len)).sum().sum()
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
        # --- Append word search results to log ---
        log_lines.append("")
        log_lines.append("Word Search Results:")
        if check_ai_bot_search:
            if word_search_serials and len(word_search_serials) > 0:
                for serial in word_search_serials:
                    log_lines.append(str(serial))
            else:
                log_lines.append("No respondents flagged by word search.")
        else:
            log_lines.append("Word search not applied.")

        with open(logfile_path, 'w', encoding='utf-8') as logf:
            for line in log_lines:
                logf.write(line + '\n')
        if status_callback:
            status_callback('Done! Files created.',)
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