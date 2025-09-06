# DataExtractQC GUI User Guide

Welcome to DataExtractQC! This guide will help you understand what the program does and how to use it, even if you have no technical background.

## What is DataExtractQC?
DataExtractQC is a simple tool that helps you extract and organize survey data from Excel files. It is designed for people who work with survey results and need to quickly create clean, organized files for analysis or reporting.

## What does it do?
- Reads your survey data from an Excel (.xlsx) file.
- Lets you choose which columns (questions) to keep, using an "include.txt" file.
- Automatically finds and adds open-ended questions (like text responses) if you want.
- Splits your data into separate files by language, plus an overall file.
- Saves all results in a folder you choose.
- Shows you clear progress and messages as it works.

## How do I use it?
1. **Open the program**
   - Double-click the DataExtractQC.exe file (or run `python data_extract_gui.py` if you use Python).

2. **Select your files and folder**
   - Click "Browse" next to **Data File (.xlsx)** and pick your Excel survey file.
   - Click "Browse" next to **Include.txt file** and pick your include file (a text file listing the columns you want).
   - Click "Browse" next to **Output folder** and pick where you want the results saved.

3. **(Optional) Automatically find open ends**
   - If you want the program to look for open-ended questions, check the box "Automatically search the data file for open ends."

4. **Run the extraction**
   - Click the **Run** button.
   - The program will show you what it's doing (e.g., loading your file, finding open ends, creating output files).
   - When finished, you'll see a message that your files are ready.

5. **Check your results**
   - Go to the output folder you chose. You'll find:
     - An overall Excel file with your selected columns.
     - Separate Excel files for each language in your data.
     - A log file with details about what was done.
     - (If you used open end detection) a new include file with open ends added.

## Tips
- If you see an error, check the error_log.txt file in your data folder for details.
- You can use the same include.txt file for different surveys, just update the column names as needed.
- The program will never overwrite your original include.txt file.

## Need help?
If you have questions or problems, contact your data team or open an issue on the GitHub page.

---

**DataExtractQC makes survey data extraction fast and easy—no technical skills required!**
