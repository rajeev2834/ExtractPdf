import os
import pandas as pd

# Folder path
folder_path = r"E:\Election\General Elections\4_Assembly Election - 2020\Electoral Roll Excel File\AC_235"

# Search text (set directly)
search_text = "Mosafir Prasad"

# Loop through all Excel files
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        file_path = os.path.join(folder_path, file)
        print(f"🔎 Searching in file: {file} ...")  # progress message
        try:
            # Read all sheets
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)  # read as string
                if df.apply(lambda col: col.str.contains(search_text, case=False, na=False)).any().any():
                    print(f"✅ Found in {file} | Sheet: {sheet}")
        except Exception as e:
            print(f"⚠️ Could not read {file}: {e}")
