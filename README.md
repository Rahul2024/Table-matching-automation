
Excel Data Reconciliation Tool
This Python script automates the comparison of two Excel sheets (Sheet1 and Sheet2) based on a specified key column. It performs fuzzy matching of column names, compares row values (handling data types like dates, integers, and strings), and generates a detailed Excel report highlighting matches, mismatches, and missing data.

Features
Automated Key Matching: Finds the specified key column (HASH_KEY by default) in both sheets, using fuzzy matching if an exact name match isn't found.
Fuzzy Column Matching: Uses rapidfuzz to intelligently match column names between the two sheets, even if names are slightly different.
Data Type Awareness:
Automatically detects columns containing "Date" in their name.
Converts numeric Excel serial dates (e.g., 44929) to proper Python date objects.
Preserves integer and float data types.
Ensures correct formatting (short date format like DD/MM/YYYY) in the output Excel file.
Robust Comparisons:
Case-insensitive string comparisons.
Specific handling for date object comparisons.
Handles leading zeros in key columns correctly.
Comprehensive Output Report: Generates an Excel file with multiple sheets:
Source Data: The original data from Sheet1.
Target Data: The original data from Sheet2.
Row Comparison: A boolean matrix indicating whether values in corresponding columns match (True) or not (False) for each row.
Overall Result: A summary showing the status (match/mismatch/missing) and KPI (PASS/FAIL) for each column pair or unmatched column.
Column Mapping: Details the fuzzy matching results between source and target columns.
Side by Side Result: Displays data from both sheets side-by-side for easy visual comparison, including key columns and matched pairs.
Mismatch Details: A filtered view showing only the rows and columns where mismatches were detected.
Performance Optimizations: Employs vectorized Pandas operations and rapidfuzz for significantly faster execution compared to row-by-row loops. Avoids DataFrame fragmentation warnings.
Clear Formatting: Applies color-coding (orange for keys, grey for target columns, green for PASS, red for FAIL/mismatches) and auto-fits column widths in the output Excel file.
Prerequisites
You need Python 3.x installed. The script requires the following Python libraries:

pandas
numpy
rapidfuzz
openpyxl
Installation
Clone the Repository:
bash


1
2
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
(Recommended) Create a Virtual Environment:
bash


1
2
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
Install Dependencies:
bash


1
pip install pandas numpy rapidfuzz openpyxl
Configuration
Before running the script, you need to configure the input file path and key column(s) in the script itself.

Open the Python script file (e.g., reconcile.py) in a text editor.
Modify the file_path variable to point to your Excel workbook:
python


1
file_path = "/path/to/your/input_file.xlsx"
Modify the key_columns list if your key column has a different name:
python


1
key_columns = ['YOUR_KEY_COLUMN_NAME'] # e.g., ['ID'], ['Order Number']
(Optional) Change the output_file path if desired:
python


1
output_file = "/path/to/your/output_file.xlsx"
(Optional) Adjust the date format in the set_excel_column_types function:
python


1
date_style.number_format = 'DD/MM/YYYY' # Change to 'MM/DD/YYYY', 'YYYY-MM-DD', etc.
Usage
Ensure you have configured the script as described above.

Run the script from your terminal:

bash


1
python reconcile.py
The script will process the data and generate the output Excel file specified by output_file. Open this file to view the reconciliation results.

Example
Given an input file data.xlsx with sheets Sheet1 and Sheet2, and a key column ID, running the script will produce data_reconciled.xlsx containing the detailed comparison report across the seven sheets described in the Features section.

Contributing
Feel free to fork the repository and submit pull requests for improvements or bug fixes.
