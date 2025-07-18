import json
import pandas as pd
import os

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink

# Read JSON from TXT
txt_path = r"C:\Users\NIC\OneDrive\Desktop\Booth List - 2003\243.txt"
output_excel = "nawada_parts_with_links_243.xlsx"

if not os.path.exists(txt_path):
    print(f"❌ File not found at: {txt_path}")
    exit(1)

with open(txt_path, "r", encoding="utf-8") as f:
    full_text = f.read()

json_start = full_text.find("{")
json_end = full_text.rfind("}") + 1
json_string = full_text[json_start:json_end]
data = json.loads(json_string)

# Build DataFrame
records = []
for item in data.get("payload", []):
    records.append({
        "acNumber": item.get("acNumber", ""),
        "partNumber": item.get("partNumber", ""),
        "partName": item.get("partName", ""),
        "Link": item.get("oldPdfUrl", "").strip().replace('"', '')
    })

df = pd.DataFrame(records)

# Create Excel Workbook
wb = Workbook()
ws = wb.active
ws.title = "Parts List"

# Write headers
ws.append(["acNumber", "partNumber", "partName", "Link"])

# Write rows with hyperlink
for _, row in df.iterrows():
    ws.append([
        row["acNumber"],
        row["partNumber"],
        row["partName"],
        f'=HYPERLINK("{row["Link"]}", "View")'
    ])

# Save to Excel
wb.save(output_excel)
print(f"✅ Excel with labeled hyperlinks saved to: {output_excel}")