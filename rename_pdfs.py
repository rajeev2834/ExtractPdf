import os
import re
import pytesseract
from pdf2image import convert_from_path

# Set Tesseract path (Update this based on your installation)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\poppler\bin'  # Update this path to your poppler bin directory

def extract_text_from_pdf(pdf_path):
    """
    Converts a scanned PDF to images and extracts text using Tesseract OCR.
    """
    try:
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        extracted_text = []

        for image in images:
            text = pytesseract.image_to_string(image)
            extracted_text.append(text)

        return "\n".join(extracted_text).strip()

    except Exception as e:
        print(f"Error extracting text from {pdf_path}: {e}")
        return None

def extract_name_from_text(text):
    """
    Extracts the name appearing after 'I,' and before 'son'.
    """
    match = re.search(r'I,\s*([^,]*?)\s*(?=son)', text)
    return match.group(1).strip() if match else None

def rename_pdf_file(pdf_path):
    """
    Renames the PDF file based on the extracted name.
    """
    text = extract_text_from_pdf(pdf_path)
    if text:
        extracted_name = extract_name_from_text(text)
        if extracted_name:
            directory, old_filename = os.path.split(pdf_path)
            new_filename = f"{extracted_name}.pdf"
            new_filepath = os.path.join(directory, new_filename)

            try:
                os.rename(pdf_path, new_filepath)
                print(f"Renamed '{old_filename}' → '{new_filename}'")
            except Exception as e:
                print(f"Error renaming {pdf_path}: {e}")
        else:
            print(f"No valid name extracted from {pdf_path}")
    else:
        print(f"Skipping {pdf_path} due to text extraction failure")

def process_pdf_directory(directory):
    """
    Processes all scanned PDFs in the given directory.
    """
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                rename_pdf_file(pdf_path)

if __name__ == "__main__":
    print('inside main')
    pdf_directory = 'C:/Users/rajee/OneDrive/Desktop/Assets Declaration'
    process_pdf_directory(pdf_directory)