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
try:
    from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
    VADER_AVAILABLE = True
except ImportError:
    VADER_AVAILABLE = False
    print("Warning: vaderSentiment not available. Sentiment analysis will be skipped.")

# Enhanced sentiment analysis with complaint detection
COMPLAINT_INDICATORS = [
    "feel deceived", "disappointed", "not the first time", 
    "longer than stated", "inaccurate", "misleading", "dishonest",
    "not what was promised", "expected better", "waste of time",
    "frustrated", "annoyed", "should have been told", "poor service",
    "twice as long", "much longer", "far longer", "way longer",
    "completely wrong", "totally wrong", "absolutely wrong",
    "very disappointed", "extremely disappointed", "highly disappointed",
    "poor quality", "terrible quality", "awful quality", "bad quality",
    "unprofessional", "unacceptable", "ridiculous", "outrageous", "too long", "disingenuous survey", "very long", "same question", "longer than estimated", "impossible to answer"
]

def detect_complaints(text):
    """
    Detect complaint patterns in text.
    Returns: (is_complaint: bool, complaint_count: int, found_patterns: list)
    """
    if not text or pd.isna(text):
        return False, 0, []
    
    text_lower = str(text).lower()
    found_patterns = []
    
    for indicator in COMPLAINT_INDICATORS:
        if indicator in text_lower:
            found_patterns.append(indicator)
    
    complaint_count = len(found_patterns)
    is_complaint = complaint_count > 0
    
    return is_complaint, complaint_count, found_patterns

def enhanced_sentiment_score(text, vader_score):
    """
    Apply weighted scoring to adjust VADER sentiment for obvious complaints.
    Returns: (adjusted_score: float, adjustment_applied: float, complaint_patterns: list)
    """
    is_complaint, complaint_count, found_patterns = detect_complaints(text)
    
    if is_complaint:
        # The more complaint indicators, the more we adjust downward
        # Each indicator reduces score by 0.3, capped at 0.8 total reduction
        adjustment = min(complaint_count * 0.3, 0.8)
        adjusted_score = vader_score - adjustment
        # Don't go below -1.0 (VADER's minimum)
        adjusted_score = max(adjusted_score, -1.0)
        return adjusted_score, adjustment, found_patterns
    
    return vader_score, 0.0, []

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

def run_data_extract(input_file, include_file, output_dir, check_open_ends=True, check_ai_bot_search=False, word_file=None, error_log_path=None, status_callback=None, check_duplicate_postcode_yob=False, check_length=True, length_multiplier=10, check_sentiment=False, sentiment_pos_threshold=0.6, sentiment_neg_threshold=0):
    print("[DEBUG] Starting run_data_extract")
    import time
    from datetime import datetime


    # Ensure output_dir is a string for os.path checks
    output_dir_str = str(output_dir)
    if not os.path.isdir(output_dir_str):
        try:
            os.makedirs(output_dir_str, exist_ok=True)
        except Exception:
            # If we can't create the output dir, fallback to input file dir
            output_dir_str = os.path.dirname(input_file)
    if error_log_path is None:
        error_log_path = os.path.join(output_dir_str, ERROR_LOG_FILENAME)

    start_time = datetime.now()
    start_ts = time.time()

    # Step 1: Load Excel file
    print("[DEBUG] Step 1: Loading Excel file...")
    try:
        if status_callback:
            status_callback(STATUS_MESSAGES['load_excel'])
        df = pd.read_excel(input_file, dtype=str)
    except Exception as e:
        log_error(f'Error reading Excel file: {e}', error_log_path)
        print("[DEBUG] Error occurred during Step 1: Loading Excel file")
        if status_callback:
            status_callback('Error: Could not read the Excel file. Please check the file and try again.')
        return False

    # Step 2: Handle include.txt and open ends
    print("[DEBUG] Step 2: Handling include.txt and open ends...")
    try:
        open_end_cols = []
        original_include_columns = []
        output_dir = Path(output_dir)
        include_with_opens_path = output_dir / INCLUDE_WITH_OPENS_FILENAME
        if check_open_ends:
            if status_callback:
                status_callback(STATUS_MESSAGES['scan_open_ends'])
            try:
                open_end_cols = get_open_ends(df)
            except Exception as e:
                log_error(f'Error in get_open_ends: {e}', error_log_path)
                if status_callback:
                    status_callback('Error: Could not extract open ends from the data.')
                return False
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
        print("[DEBUG] Error occurred during Step 2: Handling include.txt and open ends")
        if status_callback:
            status_callback('Error: Unexpected error handling include.txt.')
        return False

    # Step 3: Filter columns
    print("[DEBUG] Step 3: Filtering columns...")
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
        print("[DEBUG] Error occurred during Step 3: Filtering columns")
        if status_callback:
            status_callback('Error: Problem processing selected columns.')
        return False

    # Step 4: Check output directory
    print("[DEBUG] Step 4: Checking output directory...")
    try:
        if not os.path.isdir(output_dir):
            if status_callback:
                status_callback(f'Error: Output directory does not exist: {output_dir}')
            return False
    except Exception as e:
        log_error(f'Error during output directory prompt: {e}', error_log_path)
        print("[DEBUG] Error occurred during Step 4: Checking output directory")
        if status_callback:
            status_callback('Error: Problem with output directory.')
        return False

    # Step 5: Find unique InterviewLanguage values
    print("[DEBUG] Step 5: Scanning for unique InterviewLanguage values...")
    try:
        if status_callback:
            status_callback(STATUS_MESSAGES['scan_languages'])
        if 'InterviewLanguage' not in df.columns or df['InterviewLanguage'].dropna().empty:
            print("[DEBUG] InterviewLanguage column missing or empty, defaulting to ENG for all rows.")
            df['InterviewLanguage'] = 'ENG'
        languages = df['InterviewLanguage'].dropna().unique()
        debug(f"Languages found: {languages}")
    except Exception as e:
        log_error(f'Error finding InterviewLanguage values: {e}', error_log_path)
        print("[DEBUG] Error occurred during Step 5: Scanning for unique InterviewLanguage values")
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
        flagged_serials = set()  # Initialize outside word search block
        matches = []  # Initialize outside word search block
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
                    log_error(f"[Word Search] Error reading word file: {e}", error_log_path)
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
            except Exception as e:
                log_error(f"[Word Search] Error during word search: {e}", error_log_path)
        
        # --- Debug: Print toggle values before creating highlighted file ---
        debug(f"[Highlight File] Toggle states:")
        debug(f"  - highlighted_cells (word search): {bool(highlighted_cells)} (count: {len(highlighted_cells) if highlighted_cells else 0})")
        debug(f"  - check_sentiment: {check_sentiment}")
        debug(f"  - check_duplicate_postcode_yob: {check_duplicate_postcode_yob}")
        debug(f"  - check_length: {check_length}")
        should_create_file = highlighted_cells or check_sentiment or check_duplicate_postcode_yob or check_length
        debug(f"[Highlight File] Should create highlighted file: {should_create_file}")
        # --- Create highlighted Excel file with CHECKS column if any relevant toggle is selected ---
        if should_create_file:
                    debug(f"[Highlight File] Creating highlighted file...")
                    highlighted_file = os.path.join(output_dir, f'{base_name}__Overall_highlighted.xlsx')
                    debug(f"[Highlight File] Highlighted file path: {highlighted_file}")
                    # Load the original overall file as DataFrame
                    df_highlight = pd.read_excel(overall_file, dtype=str)
                    debug(f"[Highlight File] Loaded dataframe with {len(df_highlight)} rows and {len(df_highlight.columns)} columns")
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

                    # === STEP 1: Add all additional columns BEFORE highlighting ===
                    
                    # Insert sentiment score column after 'outro' if sentiment analysis is enabled
                    sentiment_added = False
                    if check_sentiment and VADER_AVAILABLE and 'outro' in df_highlight.columns:
                        debug(f"[Sentiment Analysis] Adding sentiment scores column after 'outro'")
                        outro_idx = df_highlight.columns.get_loc('outro') + 1
                        sentiment_scores = []
                        analyzer = SentimentIntensityAnalyzer()
                        for val in df_highlight['outro']:
                            if pd.isna(val) or str(val).strip() == "":
                                sentiment_scores.append("")
                            else:
                                vader_score = analyzer.polarity_scores(str(val))['compound']
                                enhanced_score, adjustment, patterns = enhanced_sentiment_score(str(val), vader_score)
                                sentiment_scores.append(enhanced_score)
                        df_highlight.insert(outro_idx, 'outro_sentiment_score', sentiment_scores)
                        sentiment_added = True
                        debug(f"[Sentiment Analysis] Sentiment score column added at position {outro_idx}")

                    # === STEP 2: Save the dataframe with all columns finalized ===
                    df_highlight.to_excel(highlighted_file, index=False, engine='openpyxl')
                    
                    # Apply highlighting to cells
                    from openpyxl import load_workbook
                    from openpyxl.styles import PatternFill
                    wb = load_workbook(highlighted_file)
                    ws = wb.active
                    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                    orange_fill = PatternFill(start_color='FFFFA500', end_color='FFFFA500', fill_type='solid')
                    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
                    
                    # Calculate median lengths for each column (ignoring blanks)
                    median_lengths = {}
                    for col in df_highlight.columns:
                        non_blank = df_highlight[col].dropna().astype(str)
                        non_blank = non_blank[non_blank != ""]
                        if len(non_blank) == 0:
                            median_lengths[col] = 0
                        else:
                            median_lengths[col] = non_blank.map(len).median()
                    
                    # === STEP 3: Apply highlighting with correct column indices ===
                    
                    # Highlight word search matches - recalculate column positions after all columns are added
                    if highlighted_cells:
                        debug(f"[Word Search Highlighting] Applying word search highlights to {len(highlighted_cells)} cells")
                        
                        # Create column mapping ONCE (not per cell!)
                        column_mapping = {}
                        if check_ai_bot_search:
                            try:
                                # Get original dataframe columns ONCE
                                df_original = pd.read_excel(overall_file, dtype=str)
                                for original_idx, col_name in enumerate(df_original.columns):
                                    if col_name in df_highlight.columns:
                                        final_idx = df_highlight.columns.get_loc(col_name) + 1  # +1 for Excel 1-based
                                        column_mapping[original_idx + 1] = final_idx  # +1 because highlighted_cells uses 1-based
                                debug(f"[Word Search] Created column mapping for {len(column_mapping)} columns")
                            except Exception as e:
                                debug(f"[Word Search] Error creating column mapping: {e}")
                                column_mapping = {}
                        
                        # Apply highlighting using the pre-built mapping
                        for row_idx, original_col_idx in highlighted_cells:
                            try:
                                if original_col_idx in column_mapping:
                                    final_col_idx = column_mapping[original_col_idx]
                                    ws.cell(row=row_idx, column=final_col_idx).fill = red_fill
                                    # Only debug first few to avoid spam
                                    if len([x for x in highlighted_cells]) <= 10:
                                        debug(f"[Word Search] Highlighted cell at row {row_idx}, column {final_col_idx}")
                                else:
                                    # Fallback to old method
                                    adjustment = 1  # CHECKS column
                                    if sentiment_added:
                                        adjustment += 1  # sentiment score column if added before this column
                                    ws.cell(row=row_idx, column=original_col_idx + adjustment).fill = red_fill
                            except Exception as e:
                                debug(f"[Word Search] Error highlighting cell at row {row_idx}: {e}")
                    
                    # Highlight cells > multiplier x median in orange (skip header row)
                    if check_length:
                        for i, row in enumerate(df_highlight.itertuples(index=False), start=2):
                            for j, col in enumerate(df_highlight.columns, start=1):
                                if col == "CHECKS":
                                    continue
                                val = getattr(row, col) if hasattr(row, col) else ""
                                if pd.isna(val) or val == "":
                                    continue
                                cell_len = len(str(val))
                                median_len = median_lengths.get(col, 0)
                                if median_len > 0 and cell_len > length_multiplier * median_len:
                                    ws.cell(row=i, column=j).fill = orange_fill
                                    # If not already flagged for WORDS, add/check CHECKS column for LENGTH
                                    checks_col_idx = df_highlight.columns.get_loc("CHECKS") + 1
                                    existing = ws.cell(row=i, column=checks_col_idx).value
                                    if existing:
                                        if "LENGTH" not in existing:
                                            ws.cell(row=i, column=checks_col_idx).value = f"{existing},LENGTH"
                                    else:
                                        ws.cell(row=i, column=checks_col_idx).value = "LENGTH"
                    
                    # Highlight duplicate (postcode, yob) pairs and flag CHECKS
                    duplicate_log_lines = []
                    if check_duplicate_postcode_yob:
                        try:
                            postcode_col = None
                            yob_col = None
                            for idx, col in enumerate(df_highlight.columns):
                                if col.lower() == "personal_ros_postcode":
                                    postcode_col = col
                                if col.lower() == "personal_ros_yob":
                                    yob_col = col
                            if postcode_col and yob_col:
                                # Exclude rows where postcode is blank or contains _dk01rf_ (case-insensitive), or yob is blank
                                pairs = df_highlight[[postcode_col, yob_col]].astype(str)
                                mask_valid = (
                                    (pairs[postcode_col].str.strip() != "") &
                                    (pairs[yob_col].str.strip() != "") &
                                    (~pairs[postcode_col].str.lower().str.contains("_dk01rf_"))
                                )
                                pairs_valid = pairs[mask_valid]
                                # Find all duplicate groups
                                dup_groups = pairs_valid.groupby([postcode_col, yob_col]).filter(lambda x: len(x) > 1)
                                for (postcode, yob), group in dup_groups.groupby([postcode_col, yob_col]):
                                    for idx_df in group.index:
                                        postcode_val = str(df_highlight.at[idx_df, postcode_col]) if pd.notnull(df_highlight.at[idx_df, postcode_col]) else ""
                                        yob_val = str(df_highlight.at[idx_df, yob_col]) if pd.notnull(df_highlight.at[idx_df, yob_col]) else ""
                                        # Skip blank or refused
                                        if postcode_val.strip() == "" or yob_val.strip() == "":
                                            continue
                                        if "_dk01rf_" in postcode_val.lower():
                                            continue
                                        excel_row = idx_df + 2  # DataFrame index to Excel row
                                        ws.cell(row=excel_row, column=df_highlight.columns.get_loc(postcode_col)+1).fill = yellow_fill
                                        ws.cell(row=excel_row, column=df_highlight.columns.get_loc(yob_col)+1).fill = yellow_fill
                                        # Add DUPLICATE to CHECKS
                                        checks_col_idx = df_highlight.columns.get_loc("CHECKS") + 1
                                        existing = ws.cell(row=excel_row, column=checks_col_idx).value
                                        if existing:
                                            if "DUPLICATE" not in existing:
                                                ws.cell(row=excel_row, column=checks_col_idx).value = f"{existing},DUPLICATE"
                                        else:
                                            ws.cell(row=excel_row, column=checks_col_idx).value = "DUPLICATE"
                                        # Log serial and pair
                                        serial_val = df_highlight.at[idx_df, serial_col] if serial_col else idx_df
                                        duplicate_log_lines.append(f"Serial: {serial_val} | Postcode: {postcode_val} | YOB: {yob_val}")
                        except Exception as e:
                            log_error(f"[Duplicate Check] Error: {e}", error_log_path)
                    
                    # Sentiment Analysis highlighting
                    sentiment_log_lines = []
                    if check_sentiment and VADER_AVAILABLE:
                        debug(f"[Sentiment Analysis] Starting sentiment analysis highlighting...")
                        try:
                            analyzer = SentimentIntensityAnalyzer()
                            # Define fills for sentiment highlighting
                            green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light green
                            red_fill_sentiment = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")  # Tomato red
                            white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # White
                            
                            # Only analyze 'outro' column
                            if 'outro' in df_highlight.columns:
                                debug(f"[Sentiment Analysis] Found 'outro' column, analyzing {len(df_highlight)} rows...")
                                sentiment_log_lines.append("Analyzing sentiment in 'outro' column")
                                outro_col_idx = df_highlight.columns.get_loc('outro') + 1
                                pos_count = 0
                                neg_count = 0
                                neutral_count = 0
                                for i, row in enumerate(df_highlight.itertuples(index=False), start=2):
                                    val = getattr(row, 'outro') if hasattr(row, 'outro') else ""
                                    if pd.isna(val) or val == "" or str(val).strip() == "":
                                        continue
                                    vader_score = analyzer.polarity_scores(str(val))['compound']
                                    enhanced_score, adjustment, complaint_patterns = enhanced_sentiment_score(str(val), vader_score)
                                    # Highlight positive (>=pos_threshold) green, negative (<=neg_threshold) red, else white
                                    if enhanced_score >= sentiment_pos_threshold:
                                        ws.cell(row=i, column=outro_col_idx).fill = green_fill
                                        pos_count += 1
                                    elif enhanced_score <= sentiment_neg_threshold:
                                        ws.cell(row=i, column=outro_col_idx).fill = red_fill_sentiment
                                        neg_count += 1
                                    else:
                                        ws.cell(row=i, column=outro_col_idx).fill = white_fill
                                        neutral_count += 1
                                    text_preview = str(val)[:100] + "..." if len(str(val)) > 100 else str(val)
                                    # Log with both original and enhanced scores if adjustment was made
                                    if adjustment > 0:
                                        sentiment_log_lines.append(f"Row {i}: VADER={vader_score:.3f} -> Enhanced={enhanced_score:.3f} (adj:-{adjustment:.3f}) - {text_preview}")
                                        sentiment_log_lines.append(f"  Complaint patterns: {', '.join(complaint_patterns)}")
                                    else:
                                        sentiment_log_lines.append(f"Row {i}: score={enhanced_score:.3f} - {text_preview}")
                                sentiment_log_lines.append(f"Total very positive: {pos_count}")
                                sentiment_log_lines.append(f"Total very negative: {neg_count}")
                                sentiment_log_lines.append(f"Total neutral: {neutral_count}")
                            else:
                                sentiment_log_lines.append("No 'outro' column found for sentiment analysis")
                        except Exception as e:
                            log_error(f"[Sentiment Analysis] Error: {e}", error_log_path)
                            sentiment_log_lines.append(f"Sentiment analysis failed: {e}")
                    elif check_sentiment and not VADER_AVAILABLE:
                        debug(f"[Sentiment Analysis] VADER library not available!")
                        sentiment_log_lines.append("Sentiment analysis requested but vaderSentiment library not available.")
                    elif check_sentiment:
                        debug(f"[Sentiment Analysis] Sentiment analysis requested but no 'outro' column found")
                    
                    wb.save(highlighted_file)
                    debug(f"[Highlight] Highlighted file saved successfully: {highlighted_file}")
        else:
            debug(f"[Highlight File] No relevant toggles selected - skipping highlighted file creation")

        # Create log file with processing results
        log_lines = []
        log_lines.append("Processing Results Summary")
        log_lines.append("=" * 40)
        log_lines.append("")
        log_lines.append(f"Input file: {os.path.basename(input_file)}")
        log_lines.append(f"Output directory: {output_dir}")
        log_lines.append("")
        log_lines.append("Files created:")
        for f in output_files:
            log_lines.append(f"  {os.path.basename(f)}")
        log_lines.append("")
        log_lines.append("Total number of characters in each file:")
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

        # --- Append duplicate log if any ---
        if check_duplicate_postcode_yob and 'duplicate_log_lines' in locals() and duplicate_log_lines:
            log_lines.append("")
            log_lines.append("Duplicate Postcode/YOB Pairs Found:")
            for line in duplicate_log_lines:
                log_lines.append(line)
        
        # --- Append sentiment analysis log if any ---
        if check_sentiment and 'sentiment_log_lines' in locals() and sentiment_log_lines:
            log_lines.append("")
            log_lines.append("Sentiment Analysis Results:")
            log_lines.append(f"Positive threshold: {sentiment_pos_threshold}")
            log_lines.append(f"Negative threshold: {sentiment_neg_threshold}")
            for line in sentiment_log_lines:
                log_lines.append(line)
        
        # --- Append column median lengths at the bottom ---
        log_lines.append("")
        log_lines.append("Column Median Lengths:")
        # Use the overall file for median calculation
        try:
            df_overall_for_median = pd.read_excel(overall_file, dtype=str)
            for col in df_overall_for_median.columns:
                non_blank = df_overall_for_median[col].dropna().astype(str)
                non_blank = non_blank[non_blank != ""]
                if len(non_blank) == 0:
                    median_len = 0
                else:
                    median_len = non_blank.map(len).median()
                log_lines.append(f"  {col}: {median_len:.2f}")
        except Exception as e:
            log_lines.append(f"  [Error calculating medians: {e}]")

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