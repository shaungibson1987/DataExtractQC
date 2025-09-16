# DataExtractQC GUI User Guide

Welcome to DataExtractQC! This guide will help you understand what the program does and how to use it, even if you have no technical background.

## What is DataExtractQC?
DataExtractQC is a simple tool that helps you extract and organize survey data from Excel files. It is designed for people who work with survey results and need to quickly create clean, organized files for analysis or reporting.

## What does it do?
- Reads your survey data from an Excel (.xlsx) file.
- Lets you choose which columns (questions) to keep, using an "include.txt" file (one column name per line, no semicolons).
- Automatically finds and adds open-ended questions (like text responses) if you want.
- Splits your data into separate files by language, plus an overall file.
- Saves all results in a folder you choose.
- Shows you clear progress and messages as it works.

## How do I use it?
1. **Open the program**
   - Double-click the DataExtractQC.exe file (or run `python data_extract_gui.py` if you use Python).

2. **Select your files and folder**
   - Click "Browse" next to **Data File (.xlsx)** and pick your Excel survey file.
   - Click "Browse" next to **Include.txt file** and pick your include file (a text file listing the columns you want, one column name per line).
   - Click "Browse" next to **Output folder** and pick where you want the results saved.


3. **(Optional) Enable advanced checks**
   - All advanced checks are off by default. Enable only the ones you need:
     - **Automatically search the data file for open ends**: Finds and adds open-ended/text columns to your output.
     - **Search file for specific words (words.txt)**: Flags and highlights cells containing words from a list you provide (one word per line in a .txt file). Click "Browse" to select your word list.
     - **Check for duplicate postcode/YOB pairs**: Flags and highlights rows where the combination of postcode and year of birth is duplicated.
     - **Enable length checks**: Flags and highlights cells that are unusually long compared to others in the same column. You can set the multiplier (default is 10) to control what counts as "unusually long" (e.g., 10× the median length for that column).

4. **Run the extraction**
   - Click the **Run** button.
   - The program will show you what it's doing (e.g., loading your file, running checks, creating output files).
   - When finished, you'll see a message that your files are ready.

5. **Check your results**
   - Go to the output folder you chose. You'll find:
     - An overall Excel file with your selected columns.
     - Separate Excel files for each language in your data.
     - A log file with details about what was done.
     - (If you used open end detection) a new include file with open ends added.
     - (If you enabled advanced checks) highlighted cells and a CHECKS column in the Excel output, showing which rows were flagged and why.

## Tips
- If you see an error, check the error_log.txt file in your data folder for details.
- You can use the same include.txt file for different surveys, just update the column names as needed. Each column name should be on its own line (no semicolons).
- The program will never overwrite your original include.txt file.

## Need help?
If you have questions or problems, contact shaun.gibson@yonderdatasolutions.com or open an issue on the GitHub page.

---

**DataExtractQC makes survey data extraction fast and easy—no technical skills required!**