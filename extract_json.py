import json
import csv
import os

# File path
txt_path = r"C:\Users\NIC\OneDrive\Desktop\Booth List - 2003\243- Hisua.txt"

if not os.path.exists(txt_path):
    print(f"❌ File not found at: {txt_path}")
    exit(1)

# Load the text file
try:
    with open(txt_path, "r", encoding="utf-8") as f:
        full_text = f.read()
except Exception as e:
    print("❌ Error reading TXT file:", e)
    exit(1)

# Extract JSON string from text
json_start = full_text.find("{")
json_end = full_text.rfind("}") + 1
json_string = full_text[json_start:json_end]

# Parse JSON
try:
    data = json.loads(json_string)
except json.JSONDecodeError as e:
    print("❌ Failed to parse JSON:", e)
    exit(1)

payload = data.get("payload", [])

# Write to CSV
fields = ["acNumber", "partNumber", "partName"]
with open("nawada_parts.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=fields)
    writer.writeheader()
    for item in payload:
        writer.writerow({field: item.get(field, "") for field in fields})

print("✅ Data saved to nawada_parts.csv")
