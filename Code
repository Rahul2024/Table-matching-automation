import pandas as pd
import numpy as np
from fuzzywuzzy import process
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

# === Config ===
file_path = "/Users/rahulraj/Desktop/billing23730.xlsx"
key_columns = ['HASH_KEY']
output_file = "/Users/rahulraj/Desktop/billing23730final2345.xlsx"

# === Load Sheets ===
sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

# --- Improved Column Cleaning ---
def clean_columns(df):
    """Clean column names more thoroughly."""
    cleaned_cols = []
    for col in df.columns:
        clean_col = str(col).strip().replace('\u00A0', ' ').replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        while '  ' in clean_col:
            clean_col = clean_col.replace('  ', ' ')
        cleaned_cols.append(clean_col.strip())
    return cleaned_cols

sheet1.columns = clean_columns(sheet1)
sheet2.columns = clean_columns(sheet2)

# === Improved Key Column Validation ===
def find_key_column(df_columns, target_key):
    """Find key column using exact match first, then fuzzy matching."""
    if target_key in df_columns:
        return target_key
    if df_columns:
        match, score = process.extractOne(target_key, df_columns)
        if score >= 80:
            print(f"⚠️ Warning: Using fuzzy match '{match}' for key column '{target_key}' (score: {score})")
            return match
    return None

dupe_cols = sheet1.columns[sheet1.columns.duplicated()].tolist()
if dupe_cols:
    print("❌ Duplicate columns in Sheet1:", dupe_cols)

actual_key_columns = []
for key in key_columns:
    found_key1 = find_key_column(sheet1.columns.tolist(), key)
    found_key2 = find_key_column(sheet2.columns.tolist(), key)
    if found_key1 and found_key2:
        actual_key_columns.append(key)
        if found_key1 != key: sheet1.rename(columns={found_key1: key}, inplace=True)
        if found_key2 != key: sheet2.rename(columns={found_key2: key}, inplace=True)
    else:
        print(f"❌ Error: Key column '{key}' not found!")
        raise Exception(f"Missing key column: {key}")
key_columns = actual_key_columns

# === Store Original Key Values Before Processing ===
original_keys_sheet1 = {key: sheet1[key].copy() for key in key_columns}

# === Normalize Key Columns (Handle Leading Zeros) ===
def normalize_key_value(val):
    if pd.isna(val): return str(val)
    str_val = str(val).strip()
    return str(int(str_val)) if str_val.isdigit() else str_val

for col in key_columns:
    sheet1[f'{col}_normalized'] = sheet1[col].apply(normalize_key_value)
    sheet2[f'{col}_normalized'] = sheet2[col].apply(normalize_key_value)
normalized_key_cols = [f'{col}_normalized' for col in key_columns]

# === Create Unique Composite Keys ===
sheet1['__key__'] = sheet1[normalized_key_cols].agg('|'.join, axis=1)
sheet2['__key__'] = sheet2[normalized_key_cols].agg('|'.join, axis=1)

if sheet1['__key__'].duplicated().any():
    sheet1['__key__'] += '_row' + (sheet1.index + 1).astype(str)
if sheet2['__key__'].duplicated().any():
    sheet2['__key__'] += '_row' + (sheet2.index + 1).astype(str)

sheet1.set_index('__key__', inplace=True)
sheet2.set_index('__key__', inplace=True)

composite_key_values = sheet1.index.copy()

# === Fuzzy Match Columns ===
sheet1_cols = [c for c in sheet1.columns if c not in key_columns and not c.endswith('_normalized')]
sheet2_cols = [c for c in sheet2.columns if c not in key_columns and not c.endswith('_normalized')]

matched_cols, unmatched_cols = {}, []
for col1 in sheet1_cols:
    match, score = process.extractOne(col1, sheet2_cols)
    if score >= 93:   #changed the factor by 93 from 90 for better accuracy
        matched_cols[col1] = match
    else:
        unmatched_cols.append(col1)

# === Build Side-by-side Sheet ===
sheet1_comparison_result = pd.DataFrame(index=sheet1.index)
for col1 in sheet1_cols:
    sheet1_comparison_result[col1] = sheet1[col1]
    if col1 in matched_cols:
        col2 = matched_cols[col1]
        sheet1_comparison_result[f"{col2} (target)"] = sheet2[col2].reindex(sheet1.index)
    else:
        sheet1_comparison_result[f"Missing (target)"] = "Missing in Target"

sheet1_comparison_result.insert(0, '__key__', composite_key_values)
for key in reversed(key_columns):
    sheet1_comparison_result.insert(1, key, original_keys_sheet1[key])

# --- Case-Insensitive Comparison Function ---
def are_values_equal(v1, v2):
    """Compares two values, ignoring case for strings."""
    if pd.isna(v1) and pd.isna(v2):
        return True
    if pd.notna(v1) and pd.notna(v2):
        if isinstance(v1, str) and isinstance(v2, str):
            return v1.strip().lower() == v2.strip().lower()
        return v1 == v2
    return False

# === Build Comparison (True/False) Sheet ===
comparison = pd.DataFrame(index=sheet1.index)
for col1, col2 in matched_cols.items():
    val1 = sheet1[col1]
    val2 = sheet2[col2].reindex(sheet1.index)
    comparison[col1] = [are_values_equal(val1.iloc[i], val2.iloc[i]) for i in range(len(val1))]

for col1 in unmatched_cols:
    comparison[col1] = "Missing in Sheet2"

# === Overall result Sheet (4th) ===
column_status = []
unmatched_sheet2 = set(sheet2_cols) - set(matched_cols.values())
for col2 in unmatched_sheet2: column_status.append({'Column': col2, 'Status': 'Missing in Source', 'KPI': 'FAIL'})
for col1 in unmatched_cols: column_status.append({'Column': col1, 'Status': 'Missing in Target', 'KPI': 'FAIL'})
for col in matched_cols:
    if comparison[col].dtype == bool:
        if comparison[col].all():
            column_status.append({'Column': col, 'Status': 'All values match', 'KPI': 'PASS'})
        else:
            mismatch_count = (~comparison[col]).sum()
            column_status.append({'Column': col, 'Status': f'Mismatches: {mismatch_count}/{len(comparison[col])}', 'KPI': 'FAIL'})
column_status_df = pd.DataFrame(column_status)

# === Column Mapping Sheet (5th)===
column_comparison_data = []
for col1, col2 in matched_cols.items(): column_comparison_data.append({'Source_Column': col1, 'Target_Column': col2, 'Match_Status': 'Matched', 'Fuzzy_Score': process.extractOne(col1, [col2])[1]})
for col1 in unmatched_cols: column_comparison_data.append({'Source_Column': col1, 'Target_Column': 'Not Found', 'Match_Status': 'Missing in Target', 'Fuzzy_Score': 'N/A'})
for col2 in unmatched_sheet2: column_comparison_data.append({'Source_Column': 'Not Found', 'Target_Column': col2, 'Match_Status': 'Missing in Source', 'Fuzzy_Score': 'N/A'})
column_comparison_df = pd.DataFrame(column_comparison_data)

# --- Create Mismatch Details DataFrame (Final Corrected Logic) ---
# 1. Find rows with at least one mismatch
mismatch_rows_mask = (comparison == False).any(axis=1)

# 2. Get the list of columns with a 'FAIL' KPI from the "Overall Result" sheet
failed_cols_df = column_status_df[column_status_df['KPI'] == 'FAIL']
failed_source_cols = failed_cols_df['Column'].tolist()

# 3. Build the final list of columns to show in the mismatch sheet
# We always include the key columns for context
final_mismatch_cols = ['__key__'] + key_columns

# Add the source and corresponding target columns for the failed columns
for col in failed_source_cols:
    # Ensure the column is a data column, not one marked as missing in source
    if col in matched_cols:
        final_mismatch_cols.append(col)  # Add the source column
        target_col = matched_cols[col]
        final_mismatch_cols.append(f"{target_col} (target)") # Add the target column

# 4. Filter the side-by-side results for both mismatch rows and the failed columns
mismatch_df = sheet1_comparison_result.loc[mismatch_rows_mask, final_mismatch_cols].copy()


# === Write to Excel ===
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    sheet1.reset_index().to_excel(writer, sheet_name="Source Data", index=False)
    sheet2.reset_index().to_excel(writer, sheet_name="Target Data", index=False)
    comparison.reset_index().to_excel(writer, sheet_name="Row Comparison", index=False)
    column_status_df.to_excel(writer, sheet_name="Overall Result", index=False)
    column_comparison_df.to_excel(writer, sheet_name="Column Mapping", index=False)
    sheet1_comparison_result.reset_index(drop=True).to_excel(writer, sheet_name="Side by Side Result", index=False)
    mismatch_df.reset_index(drop=True).to_excel(writer, sheet_name="Mismatch Details", index=False)

# === Apply color formatting ===
wb = load_workbook(output_file)
orange_fill = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

def auto_fit_columns_simple(worksheet):
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value or "")) for cell in column_cells)
        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 4

sheet_names = wb.sheetnames
for sheet_name in sheet_names:
    auto_fit_columns_simple(wb[sheet_name])

# Color Side by Side Result
ws_sbs = wb["Side by Side Result"]
for col_idx in range(1, ws_sbs.max_column + 1):
    header = ws_sbs.cell(row=1, column=col_idx).value
    if not header: continue
    fill = None
    if header in key_columns or header == '__key__': fill = orange_fill
    elif header.endswith("(target)"): fill = grey_fill
    if fill:
        for row_num in range(1, ws_sbs.max_row + 1):
            ws_sbs[f"{get_column_letter(col_idx)}{row_num}"].fill = fill

# Color Overall Result
ws_or = wb["Overall Result"]
for row in ws_or.iter_rows(min_row=2):
    kpi = row[2].value
    fill = green_fill if kpi == "PASS" else red_fill if kpi == "FAIL" else None
    if fill:
        for cell in row: cell.fill = fill

# Color Column Mapping
ws_cm = wb["Column Mapping"]
for row in ws_cm.iter_rows(min_row=2):
    status = row[2].value
    fill = green_fill if status == "Matched" else red_fill if "Missing" in status else None
    if fill:
        for cell in row: cell.fill = fill

# Color Mismatch Details Sheet
if "Mismatch Details" in wb.sheetnames:
    ws_md = wb["Mismatch Details"]
    headers = [cell.value for cell in ws_md[1]]
    
    mismatch_comparison_df = comparison[mismatch_rows_mask].copy()

    for r_idx, row in enumerate(ws_md.iter_rows(min_row=2), start=2):
        row_key = row[0].value
        for c_idx, cell in enumerate(row, start=1):
            header = headers[c_idx - 1]
            original_col_name = header.replace(" (target)", "")

            if original_col_name in mismatch_comparison_df.columns:
                is_match = mismatch_comparison_df.loc[row_key, original_col_name]
                if is_match == False:
                    cell.fill = red_fill

wb.save(output_file)
print(key_columns)
print(f"✅ All done! Thank you Rahul '{output_file}'.")
