# DataExtractQC GUI

A user-friendly tool for extracting and processing survey data from Excel files, with both CLI and GUI interfaces.

## Features
- Simple graphical interface (Tkinter)
- Automatic detection of open-ended columns
- Customizable include file handling
- Detailed logging and error reporting
- Output split by language and overall
- Customizable include file handling (one column per line, no semicolons)

## Requirements
- Python 3.8+
- pandas, tqdm, openpyxl, tkinter
- (For .exe) PyInstaller

## Installation
1. Clone this repository:
   ```bash
   git clone https://github.com/shaungibson1987/DataExtractQC.git
   cd DataExtractQC
   ```
2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/Scripts/activate  # On Windows (bash)
   # or
   venv\Scripts\activate.bat     # On Windows (cmd)
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the GUI App
1. Start the GUI:
   ```bash
   python data_extract_gui.py
   ```
2. Use the interface to select:
   - Data file (.xlsx)
   - Include.txt file
   - Output folder
   - (Optional) Enable open end detection
3. Click **Run**. Progress and results will be shown in the app.

## Building the .exe
1. Make sure your virtual environment is activated.
2. Install PyInstaller if needed:
   ```bash
   pip install pyinstaller
   ```
3. Build the executable:
   ```bash
   pyinstaller --onefile --windowed data_extract_gui.py
   ```
4. The .exe will be in the `dist` folder.

## Troubleshooting
- Errors are logged to `error_log.txt` in the data file's folder.
- For help, check the log or open an issue on GitHub.

## License
MIT
