**Excel Sheet Comparison & Reconciliation Tool**

An automated Python script to intelligently compare two sheets within an Excel file. The tool identifies differences in data on a row-by-row and column-by-column basis, handling messy data, and generates a detailed, color-coded report highlighting all discrepancies.

It's designed to replace tedious manual VLOOKUPs and "diff" checks with a fast, accurate, and resilient automated process.

✨ **Key Features**

Intelligent Column Matching: Uses fuzzy string matching (fuzzywuzzy) to automatically pair up columns between the two sheets, even if the headers aren't identical (e.g., "Profit Center" vs. "Profit Center(Transaction Data)").

Robust Key Normalization: Before comparing, it cleans and standardizes the key columns to ensure accurate matching. It handles:

Leading/trailing whitespace.

Numeric values stored as text (e.g., '00123' vs 123).

Numbers with trailing .0 (e.g., 456.0 vs 456).

Composite Key Support: Allows you to define multiple columns that together form a unique key for each row.

Smart Value Comparison: The comparison logic is not a simple A == B. It intelligently treats various "empty-like" values as equivalent (e.g., 0, *, #, blank cells, None, NaN).

Comprehensive Reporting: Generates a multi-sheet Excel workbook with a full breakdown of the comparison results.

Visual Formatting: The output Excel file is beautifully color-coded for immediate visual feedback:

🟢 Green for matching data and passed columns.

🔴 Red for mismatches and failed columns.

🟠 Orange for key columns.

⚪ Grey for target columns in the side-by-side view.

📋 **Output Explained**

The script generates an Excel file with the following sheets, providing a 360-degree view of the comparison:

Source Data: An exact copy of your first sheet, with the internal composite key added for reference.

Target Data: An exact copy of your second sheet, with the internal composite key added for reference.

Side by Side Result: A consolidated view showing data from the source sheet next to the corresponding data from the target sheet for every matched column.

Mismatch Details: This is the most actionable sheet. It's a filtered version of the "Side by Side Result" that only shows rows containing at least one data mismatch. Furthermore, it only includes the key columns and the specific columns where the data did not match, highlighting the differing cells in red.

Row Comparison: A technical boolean (TRUE/FALSE) report. Each cell indicates whether the value in that row and column matched its counterpart in the other sheet.

Overall Result: A high-level summary report card. It lists every column and provides a PASS / FAIL status, showing which columns match perfectly and which have mismatches or are missing.

Column Mapping: A transparency report showing how the script matched columns from the source sheet to the target sheet, including the fuzzy match score. It also lists columns that were present in one sheet but not the other.

⚙️ **Setup and Usage**

Follow these steps to get the tool running on your machine.

1. Prerequisites
Python 3.7 or newer.

The pip package manager.

2. Installation
Clone this repository or download the script to your local machine.

Navigate to the project directory in your terminal.

# Create virtual environment (optional but good practice)
python -m venv myenv

# Activate it (Windows)
myenv\Scripts\activate

# Activate it (Mac/Linux)
source myenv/bin/activate

# Then install packages
pip install pandas numpy fuzzywuzzy openpyxl python-levenshtein


# === Config ===
# The full path to your input Excel file.
file_path = "/path/to/your/input_data.xlsx"

# The column headers that uniquely identify a row.
# The script will try to find these using exact and fuzzy matching.
key_columns = ['Company Code', 'Profit Center(Transaction Data)', 'Billing Document']

# The full path where the output report will be saved.
output_file = "/path/to/your/desktop/reconciliation_report.xlsx"
file_path: Update this with the location of the Excel file you want to analyze. Your file should have two sheets, named Sheet1 and Sheet2 by default (or you can modify the pd.read_excel lines to match your sheet names).

key_columns: This is the most important setting. List the column names that together form a unique identifier for each row (a composite key).

output_file: Specify the desired name and location for the generated report.

4. Run the Script
Once configured, simply run the script from your terminal:

Bash

python your_script_name.py
You'll see progress messages in the console, and upon completion, the message ✅ All done! Results saved to '...' will appear. Your detailed, color-coded Excel report will be ready at the output path you specified.
