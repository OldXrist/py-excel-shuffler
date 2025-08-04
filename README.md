# Excel Row Shuffler

A Python script to shuffle rows within a specified range in an Excel (.xlsx) file while preserving cell formatting. Row heights are set to auto-fit content when the output file is opened in Excel.

### Requirements:
- Python 3.11 or higher
- openpyxl library (install with: pip install openpyxl)

### Directory Structure:
- input/: Place your input .xlsx files here.
- output/: Shuffled files are saved here with _shuffled appended to the file name.

### Usage:
1. Place your Excel file in the input folder.
2. Run the script from the command line:
   python shuffle_excel_preserve_format.py
3. Follow the prompts to enter:
   - Excel file name (e.g., data.xlsx)
   - Row range to shuffle (e.g., 1-10)
   - Column range to shuffle (e.g., 1-5)
4. The shuffled file will be saved in the output folder (e.g., data_shuffled.xlsx).

### Notes:
- Only .xlsx files are supported.
- Cell formatting (fonts, borders, fills, etc.) is preserved.
- Row heights auto-adjust to fit content when opened in Excel.
- Ensure the input and output folders exist (created automatically if not).