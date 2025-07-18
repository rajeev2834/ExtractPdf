import re
import camelot
import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np
import pandas as pd
from PIL import Image

# Path setup
pdf_path = r"C:\Users\NIC\OneDrive\Desktop\Booth List - 2003\239\38 Nawada_A239_A2390001.pdf"
poppler_path = r"C:\poppler-24.08.0\poppler-24.08.0\Library\bin"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Read table structure
tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
pages = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)

scale = 300 / 72  # because camelot uses 72 dpi

def fix_cid_text(text):
    if not isinstance(text, str):
        return text
    # Replace known bad CIDs
    text = text.replace("(cid:6)", "ि")
    text = text.replace("(cid:12)", "ि")
    text = text.replace("(cid:19)", "ि")
    text = re.sub(r"\(cid:\d+\)", "", text)  # remove any other leftover
    text = re.sub(r"\s+", " ", text)  # cleanup multiple spaces
    return text.strip()

# OpenCV preprocessing
def preprocess_for_ocr(pil_image):
    cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
    gray = cv2.equalizeHist(gray)
    gray = cv2.GaussianBlur(gray, (3,3), 0)
    _, threshed = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(threshed)

# Prepare final table
final_rows = []

# Process each table
for table in tables:
    df = table.df
    page_num = table.page - 1
    page_image = pages[page_num]
    height = page_image.height

    for row_idx, row in df.iterrows():
        new_row = []
        for col_idx, cell in enumerate(row):
            # Only OCR on likely Hindi name columns
            if col_idx in [5,6,8,9]:  # adjust based on your actual table
                try:
                    bbox = table._bbox[row_idx][col_idx]
                    x1, y1, x2, y2 = [int(coord * scale) for coord in bbox]
                    y1_flipped = height - y2
                    y2_flipped = height - y1

                    cropped = page_image.crop((x1, y1_flipped, x2, y2_flipped))
                    processed = preprocess_for_ocr(cropped)

                    text = pytesseract.image_to_string(
                        processed, 
                        lang='hin', 
                        config='--psm 6 -c preserve_interword_spaces=1'
                    )
                    cleaned = hindi_autocorrect(text.replace('\n', ' '))
                    new_row.append(cleaned)
                except Exception as e:
                    new_row.append(cell)
            else:
                new_row.append(cell)
        final_rows.append(new_row)

# Save to Excel
columns = tables[0].df.columns if tables else []
df_final = pd.DataFrame(final_rows, columns=columns)
df_final.to_excel("advanced_hindi_cleaned.xlsx", index=False)
print("🎉 Extraction done and saved to advanced_hindi_cleaned.xlsx")