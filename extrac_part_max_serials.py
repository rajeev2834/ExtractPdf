import pdfplumber
import pandas as pd
import os
from collections import defaultdict

# CONFIG
PDF_FOLDER = r"C:\Users\NIC\OneDrive\Desktop\Booth List - 2003\241"  # <-- your folder path
OUTPUT_CSV = "part_max_serials_241.csv"
SAVE_EVERY = 50  # Save every 50 files

# Storage
part_max_serial = defaultdict(int)

# Get all PDFs
pdf_files = sorted([f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")])
total_files = len(pdf_files)

print(f"📂 Starting batch processing of {total_files} PDF files...\n")

# Loop through files
for index, file_name in enumerate(pdf_files, start=1):
    file_path = os.path.join(PDF_FOLDER, file_name)

    print(f"🔄 [{index}/{total_files}] Processing file: {file_name}")

    # Try to extract part number from filename
    try:
        part_no = int(''.join(filter(str.isdigit, file_name)))
    except ValueError:
        print(f"⚠️ Skipped (invalid part number): {file_name}")
        continue

    try:
        with pdfplumber.open(file_path) as pdf:
            max_serial = 0
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if len(row) >= 3:
                            try:
                                serial_no = int(row[2])
                                max_serial = max(max_serial, serial_no)
                            except (ValueError, TypeError):
                                continue
            part_max_serial[part_no] = max_serial
            print(f"✅ Processed part {part_no} → Max serial: {max_serial}")

    except Exception as e:
        print(f"❌ Error reading {file_name}: {e}")

    # Batch save
    if index % SAVE_EVERY == 0 or index == total_files:
        print(f"💾 Saving progress... ({index}/{total_files})")
        try:
            df = pd.DataFrame([
                {"partNumber": p, "maxSerialNumberInPart": s}
                for p, s in sorted(part_max_serial.items())
            ])
            df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8")
            print("✅ Partial CSV saved.\n")
        except Exception as e:
            print(f"⚠️ Failed to save CSV after file {index}: {e}")

print("🎉 All done! Final CSV saved to:", OUTPUT_CSV)