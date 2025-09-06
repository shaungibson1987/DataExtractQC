# Data Extract QC Application

This tool helps you extract and split survey data from an Excel (.xlsx) file by language, keep only the columns you want, and save results for translation. It also creates a summary log file for easy reporting and logs any errors for troubleshooting.

## How to Use

### 1. Prepare Your Files
- Place your Excel data file (.xlsx) and your `include.txt` file (listing columns to keep, separated by `;`) somewhere on your computer.
- Example `include.txt`:
  ```
  Respondent.Serial;Resp_id;InterviewLanguage;outro;REQCQ1;REQCQ2;
  ```

### 2. Run the Application
- Double-click `data_extract.exe` (in the `dist` folder) or run it from a command prompt:
  ```
  python data_extract.py
  ```
- When prompted, enter:
  - The full path to your Excel file (e.g., `C:/Users/yourname/Documents/survey.xlsx`)
  - Whether you want to check for open ends to add to the include file (Y/N)
  - The full path to your `include.txt` file (e.g., `C:/Users/yourname/Documents/include.txt`)
  - The full path to the folder where you want to save the output files (e.g., `C:/Users/yourname/Documents/Output`)

### 3. What Happens Next
- The app will:
  - Read your Excel file and keep only the columns listed in `include.txt`.
  - Optionally scan for open ends (columns after `TESTJUMP` or `TESTLANG` that look like open text) and add them to `include.txt` (deduplicated).
  - Split the data by the `InterviewLanguage` column.
  - Save a separate Excel file for each language, plus an overall file, in your chosen output folder.
  - Create a log file summarizing columns kept, number of rows, and character counts per country.
  - Log any errors to `error_log.txt` in the same folder as your input file.

### 4. Output Files
- You’ll find your results in the output folder you specified:
  - `YourFile__Overall.xlsx` (all data)
  - `YourFile__ENG.xlsx`, `YourFile__FRA.xlsx`, etc. (one per language)
  - `YourFile_log.txt` (summary log)
- If any error occurs, see `error_log.txt` in the same folder as your input file for details.

### 5. Troubleshooting
- **Permission Denied:** Make sure output files are not open in Excel when running the app.
- **File Not Found:** Double-check the paths you enter for your input and include files.
- **Missing Columns:** Make sure column names in `include.txt` match those in your Excel file.
- **Error Log:** If the app fails, check `error_log.txt` for details.

### 6. Requirements
- Only `.xlsx` files are supported as input.
- No need to install Python or any packages—just use the `.exe` file, or run the script with Python 3.7+ and pandas installed.

### 7. Support
If you have questions or need help, contact shaun.gibson@yonderdatasolutions.com to discuss