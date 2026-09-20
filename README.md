# SAP Cleaner

## Overview

`SAP_Cleaner.py` is a lightweight Python utility designed to clean and standardize raw SAP export files before they are used in downstream analysis, reporting, or Excel-based workflows.

The script reads an Excel or CSV file, removes empty rows and columns, normalizes header names, and exports the cleaned result as a new Excel file in the same folder as the source file.

This is especially useful when SAP exports contain inconsistent formatting, empty sections, multilingual characters, extra header rows, or messy column names that need to be standardized for further processing.

---

## What the script does

The workflow is straightforward:

1. Prompts the user for the folder path and file name.
2. Opens the Excel or CSV file.
3. Cleans the column names by removing accents and special characters.
4. Removes empty rows and columns.
5. Detects the header row and removes extra header-like rows.
6. Produces a cleaned DataFrame.
7. Saves the result as a new Excel file.

---

## Main functions

### `unicode_utf8(string)`

This helper converts accented characters to their closest ASCII equivalents.

Examples:

- `á` -> `a`
- `ñ` -> `n`
- `ç` -> `c`
- `ü` -> `u`

It also strips common special characters such as `º`.

This is useful when Excel or SAP exports include non-ASCII characters that could cause issues in downstream scripts or spreadsheet tools.

### `df_creator_str(file_path, encoding_type, sheet_in=0, sep_in=";", header_in=0)`

This function reads the input file based on its extension:

- `.xlsx`, `.xlsm`, `.xls`, `.xlsb`
- `.csv`

It normalizes column names by:

- converting them to strings
- removing accents
- removing periods
- replacing spaces with underscores

It also ensures empty cells become empty strings.

### `SAP_normalize_process(df, skip_rows=0, skip_cols=0, header_row=0)`

This is the core cleaning routine. It:

- removes rows and columns before the relevant section
- removes fully empty rows and columns
- sets the header row as the actual column names
- discards any repeated header rows
- drops remaining empty columns
- fills missing values with empty strings

### `ask_input(message, default=None, show_default=False)`

This small helper prompts the user in the terminal and returns a value, with an optional default fallback.

---

## User interaction

When the script runs, it asks for:

- the folder path
- the file name
- the sheet name or default first sheet
- the row where headers begin
- the column where headers begin

This makes it flexible for files exported from different SAP layouts or Excel structures.

---

## Inputs expected

The script is designed for files that contain tabular data with a header row, such as:

- SAP export sheets
- Excel reports
- CSV files with semicolon or similar separators

Typical fields may include:

- material codes
- descriptions
- document references
- dates
- quantities
- cost or inventory values

---

## Output generated

The cleaned result is saved as:

`<original_file_name>_clean.xlsx`

This file is created in the same directory as the original input file.

The output contains the standardized version of the dataset with cleaned headers and empty rows removed.

---

## Example usage

Run the script and enter values like:

- Folder path: `C:/Users/yourname/Desktop/data`
- File name: `SAP_export.xlsx`
- Sheet: `Sheet1`
- Header row: `4`
- Header column: `A`

The script will produce a cleaned Excel file ready for further processing.

---

## Dependencies

This project uses:

- `pandas`
- `numpy`
- `openpyxl`
- `re`

---

## Why this script is useful

SAP exports are often messy because they include:

- blank rows
- repeated headers
- inconsistent naming
- special characters
- non-standard column positions

This script standardizes the raw file so it can be used more safely in analytics, dashboards, validation rules, or additional Python automation.

---

## Summary

`SAP_Cleaner.py` is a practical data-cleaning utility for normalizing SAP-generated files into a clean tabular structure ready for analysis or export. It is especially valuable when working with messy exported Excel or CSV data that must be made consistent before further use.
