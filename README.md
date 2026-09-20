# SAP Cleaner

Small Python script to clean SAP export files before using them in Excel or analysis workflows.

## What it does

- reads an Excel or CSV file
- removes empty rows and columns
- cleans column names (accents, spaces, dots, etc.)
- detects the real header row
- exports a cleaned Excel file in the same folder

## Main idea

SAP exports often contain blank sections, repeated headers, and messy formatting. This script standardizes the file so it can be processed more easily.

## Requirements

- Python
- pandas
- numpy
- openpyxl

## Usage

Run the script and provide:

- folder path
- file name
- sheet name (or default first sheet)
- header row number
- header column letter

The output will be saved as:

`<original_file_name>_clean.xlsx`

## Example

`SAP_export.xlsx` -> `SAP_export_clean.xlsx`
