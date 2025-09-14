def log_average_column_lengths(df):
    """Log the median character length of each column (ignoring blank cells) to the debug window."""
    debug("[Column Lengths] Median character length per column (ignoring blanks):")
    for col in df.columns:
        # Drop blanks/NaN, convert to string, and measure length
        non_blank = df[col].dropna().astype(str)
        non_blank = non_blank[non_blank != ""]
        if len(non_blank) == 0:
            median_len = 0
        else:
            median_len = non_blank.map(len).median()
        debug(f"{col} - {median_len:.2f}")
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

def run_data_extract(input_file, include_file, output_dir, check_open_ends=True, check_ai_bot_search=False, word_file=None, error_log_path=None, status_callback=None):
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

        # --- Word search post-processing ---
        if status_callback and check_ai_bot_search:
            status_callback(STATUS_MESSAGES['word_search'])
        word_search_serials = None
        highlighted_cells = set()
        if check_ai_bot_search:
            import re
            from openpyxl import load_workbook
            from openpyxl.styles import PatternFill
            # Read search words from file
            search_words = []
            if word_file:
                try:
                    with open(word_file, 'r', encoding='utf-8') as wf:
                        search_words = [line.strip() for line in wf if line.strip()]
                except Exception as e:
                    debug(f"[Word Search] Error reading word file: {e}")
            if not search_words:
                search_words = ["AI", "CHATBOT", "BOT"]  # fallback
            debug(f"[Word Search] Search words: {search_words}")
            try:
                df_overall = pd.read_excel(overall_file, dtype=str)
                # Log first 5 outro values
                if 'outro' in df_overall.columns:
                    outro_vals = df_overall['outro'].dropna().astype(str).head(5).tolist()
                    debug(f"[Word Search] First 5 outro values: {outro_vals}")
                else:
                    debug("[Word Search] No 'outro' column found in overall file.")
                matches = []
                flagged_serials = set()
                search_words_lower = [w.lower() for w in search_words]
                for idx, row in df_overall.iterrows():
                    interview_lang = str(row.get('InterviewLanguage', '')).strip().upper()
                    for col in df_overall.columns:
                        cell = str(row[col]) if not pd.isna(row[col]) else ""
                        # Tokenize using Unicode word characters
                        tokens = re.findall(r'\w+', cell, flags=re.UNICODE)
                        tokens_lower = [t.lower() for t in tokens]
                        # Remove 'ai' from tokens_lower for Italian rows
                        if interview_lang in ("ITA", "ITALIAN"):
                            tokens_lower = [t for t in tokens_lower if t != "ai"]
                        for word in search_words_lower:
                            if word in tokens_lower:
                                respondent_serial = row.get('serial', idx)
                                matches.append((respondent_serial, col, cell))
                                flagged_serials.add(respondent_serial)
                                # Track cell for highlighting: (row_idx, col_idx)
                                highlighted_cells.add((idx + 2, df_overall.columns.get_loc(col) + 1))  # +2 for header and 1-based index
                                break  # Only log once per cell
                word_search_serials = sorted(flagged_serials, key=lambda x: str(x))
                if matches:
                    debug(f"[Word Search] Matches found:")
                    for respondent_serial, col, cell in matches:
                        debug(f"  Respondent: {respondent_serial}, Column: {col}, Value: {cell}")
                else:
                    debug("[Word Search] No matches found.")
                # --- Create highlighted Excel file with CHECKS column if matches found ---
                if highlighted_cells:
                    highlighted_file = os.path.join(output_dir, f'{base_name}__Overall_highlighted.xlsx')
                    # Load the original overall file as DataFrame
                    df_highlight = pd.read_excel(overall_file, dtype=str)
                    # Find index of respondent.serial column
                    serial_col = None
                    for i, col in enumerate(df_highlight.columns):
                        if col.lower() in ("serial", "respondent.serial", "respondent_serial"):
                            serial_col = col
                            break
                    # Build set of row indices with word matches
                    rows_with_match = set([idx for idx, row in df_highlight.iterrows() if row.get('serial', idx) in flagged_serials])
                    # Insert CHECKS column after serial
                    insert_at = 1
                    if serial_col and serial_col in df_highlight.columns:
                        insert_at = df_highlight.columns.get_loc(serial_col) + 1
                    checks_col = ["WORDS" if idx in rows_with_match else "" for idx in df_highlight.index]
                    df_highlight.insert(insert_at, "CHECKS", checks_col)
                    # Sort so rows with 'WORDS' in CHECKS are at the top
                    # Do not sort by CHECKS; keep original order for correct highlighting
                    df_highlight.to_excel(highlighted_file, index=False, engine='openpyxl')
                    # Now highlight the cells
                    wb = load_workbook(highlighted_file)
                    ws = wb.active
                    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                    orange_fill = PatternFill(start_color='FFFFA500', end_color='FFFFA500', fill_type='solid')
                    # Calculate median lengths for each column (ignoring blanks)
                    median_lengths = {}
                    for col in df_highlight.columns:
                        non_blank = df_highlight[col].dropna().astype(str)
                        non_blank = non_blank[non_blank != ""]
                        if len(non_blank) == 0:
                            median_lengths[col] = 0
                        else:
                            median_lengths[col] = non_blank.map(len).median()
                    for row_idx, col_idx in highlighted_cells:
                        ws.cell(row=row_idx, column=col_idx + 1).fill = red_fill  # +1 to account for CHECKS column
                    # Highlight cells >2.5x median in orange (skip header row)
                    for i, row in enumerate(df_highlight.itertuples(index=False), start=2):
                        for j, col in enumerate(df_highlight.columns, start=1):
                            if col == "CHECKS":
                                continue
                            val = getattr(row, col) if hasattr(row, col) else ""
                            if pd.isna(val) or val == "":
                                continue
                            try:
                                cell_len = len(str(val))
                                median_len = median_lengths.get(col, 0)
                                if median_len > 0 and cell_len > 10 * median_len:
                                    ws.cell(row=i, column=j).fill = orange_fill
                                    # If not already flagged for WORDS, add/check CHECKS column for LENGTH
                                    checks_col_idx = df_highlight.columns.get_loc("CHECKS") + 1
                                    existing = ws.cell(row=i, column=checks_col_idx).value
                                    if existing:
                                        if "LENGTH" not in existing:
                                            ws.cell(row=i, column=checks_col_idx).value = f"{existing},LENGTH"
                                    else:
                                        ws.cell(row=i, column=checks_col_idx).value = "LENGTH"
                            except Exception:
                                continue
                    wb.save(highlighted_file)
                    debug(f"[Word Search] Highlighted file with CHECKS column created: {highlighted_file}")
                # Log average character length per column after word search
                if status_callback:
                    status_callback(STATUS_MESSAGES['length_checks'])
                log_average_column_lengths(df_highlight)
            except Exception as e:
                debug(f"[Word Search] Error during search: {e}")
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
            log_lines.append(f"Words searched for: {', '.join(search_words)}")
            if word_search_serials and len(word_search_serials) > 0:
                # For each flagged respondent, find the first match and print serial - cell text
                serial_to_cell = {}
                for respondent_serial, col, cell in matches:
                    if respondent_serial not in serial_to_cell:
                        serial_to_cell[respondent_serial] = cell
                for serial in word_search_serials:
                    cell_text = serial_to_cell.get(serial, "")
                    log_lines.append(f"{serial} - {cell_text}")
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