import pandas as pd
import numpy as np
import re
import ntpath
import os.path
import openpyxl

#Auxilizary function designed to return the closest ascii character to the str character found
def unicode_utf8(string):
    #If the input string is not str, it is casted as str object
    if type(string) is not str:
        string = str(string, encoding='utf-8')
    
    #str characters are transformed into the closest utf8 character
    string = re.sub(u"[àáâãäå]", 'a', string)
    string = re.sub(u"[èéêë]", 'e', string)
    string = re.sub(u"[ìíîï]", 'i', string)
    string = re.sub(u"[òóôõö]", 'o', string)
    string = re.sub(u"[ùúûü]", 'u', string)
    string = re.sub(u"[ýÿ]", 'y', string)
    string = re.sub(u"[ñ]", 'n', string)
    string = re.sub(u"[ç]", 'c', string)
    string = re.sub(u"[ÀÁÂÃÄÅ]", 'A', string)
    string = re.sub(u"[ÈÉÊË]", 'E', string)
    string = re.sub(u"[ÌÍÎÏ]", 'I', string)
    string = re.sub(u"[ÒÓÔÕÖ]", 'O', string)
    string = re.sub(u"[ÙÚÛÜ]", 'U', string)
    string = re.sub(u"[Ý]", 'Y', string)
    string = re.sub(u"[Ñ]", 'N', string)
    string = re.sub(u"[Ç]", 'C', string)
    string = re.sub(u"[º]", '', string)

    #The final string is encoded in utf8 in order to avoid problems
    return string

def df_creator_str(file_path, encoding_type, sheet_in = 0, sep_in = ";", header_in=0):
    print("Reading " + ntpath.basename(file_path))
    file_type = os.path.splitext(file_path)[1].lower()
    if file_type == ".xls" or file_type == ".xlsx" or file_type == ".xlsm" or file_type == ".xlsb":
        if file_type == ".xls":
            df_aux = pd.read_excel(file_path, sheet_name=sheet_in, header=header_in, dtype=str).fillna('')
        elif file_type == ".xlsx" or file_type == ".xlsm":
            df_aux = pd.read_excel(file_path, sheet_name=sheet_in, header=header_in, engine='openpyxl', dtype=str).fillna('')
        elif file_type == ".xlsb":
            df_aux = pd.read_excel(file_path, sheet_name=sheet_in, header=header_in, engine='pyxlsb', dtype=str).fillna('')
    elif file_type == ".csv":
        #df_aux = pd.read_csv(file_path, sep=sep_in, skiprows = header_in, error_bad_lines=False)
        print(encoding_type)
        df_aux = pd.read_csv(file_path, sep=sep_in, skiprows = header_in, on_bad_lines='skip', encoding = encoding_type, dtype=str).fillna('')
    cols = df_aux.columns
    cols = cols.map(lambda x: unicode_utf8(x) if isinstance(x, str) else x)
    cols = cols.map(lambda x: x.replace('.', '') if isinstance(x, str) else x)
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, str) else x)
    df_aux.columns = cols
    return df_aux

def SAP_normalize_process (
    df: pd.DataFrame,
    skip_rows: int = 0,
    skip_cols: int = 0,
    header_row: int = 0
) -> pd.DataFrame:

    df = df.iloc[skip_rows:, skip_cols:]

    df = df.replace('', np.nan)
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    
    df.columns = df.iloc[header_row].astype(str)
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    header = df.columns.tolist()
    non_empty_cols = [i for i, h in enumerate(header) if str(h).strip() not in ('', 'nan')]
    mask_header_rows = (df.iloc[:, non_empty_cols].astype(str) == [header[i] for i in non_empty_cols]).all(axis=1)
    df = df[~mask_header_rows].reset_index(drop=True)
    
    df = df.dropna(axis=1, how='all')
    df = df.fillna('')

    return df

def ask_input (message, default=None, show_default=False):

    if default is not None and show_default:
        prompt = f"{message} [default: {default}]:"
    else:
        prompt = f"{message}: "
    
    response = input(prompt).strip()
    return response if response else default




main_path = ask_input ("Enter the folder path")

file_name = ask_input ("Enter the name of the file with extension (.xlsx, .xlsm, ...)")

sheet_raw = ask_input ("Enter the name of the file sheet or 'ENTER' for the first appeared", default=0)

row_raw = ask_input ("Enter the row where the headers starts (e.g. 4) or 'ENTER' for default(1)", default=0)
if row_raw == 0:
    skiprows = int(row_raw)
elif row_raw == 1:
    skiprows = int(row_raw) - 1
else:
    skiprows = int(row_raw) - 2
col_raw = ask_input ("Enter the column where the headers starts (e.g. AH) or 'ENTER' for default(A)", default='A')
skipcols = openpyxl.utils.column_index_from_string(col_raw.strip().upper()) - 1



path = os.path.join(main_path, file_name)
output_file = os.path.join(main_path, f"{os.path.splitext(file_name)[0]}_clean.xlsx")
df = df_creator_str(path, "latin1", sheet_in=sheet_raw)
df = SAP_normalize_process (df, skip_rows=skiprows, skip_cols=skipcols)

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)

print(f"\nProcess completed: File {file_name} in same folder as input file")