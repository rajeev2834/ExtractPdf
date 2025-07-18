import os
import win32com.client

def add_borders_to_used_range(sheet):
    # Excel border constants
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12
    xlContinuous = 1
    xlThin = 2

    used_range = sheet.UsedRange
    for border_id in [xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal]:
        border = used_range.Borders(border_id)
        border.LineStyle = xlContinuous
        border.Weight = xlThin

def convert_excel_to_pdf_with_borders(source_folder, target_folder):
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    for file_name in os.listdir(source_folder):
        if file_name.endswith(('.xls', '.xlsx', '.xlsm')):
            excel_path = os.path.join(source_folder, file_name)
            workbook = excel.Workbooks.Open(excel_path)

            try:
                for sheet in workbook.Worksheets:
                    add_borders_to_used_range(sheet)

                pdf_file_name = os.path.splitext(file_name)[0] + ".pdf"
                pdf_path = os.path.join(target_folder, pdf_file_name)

                workbook.ExportAsFixedFormat(0, pdf_path)
                print(f"Converted with borders: {file_name} -> {pdf_file_name}")
            except Exception as e:
                print(f"Error processing {file_name}: {e}")
            finally:
                workbook.Close(False)

    excel.Quit()
    print("✅ All files processed and saved as PDFs with borders.")

# Example usage
source_folder = r"C:\Users\NIC\OneDrive\Desktop\Election\Source\236"
target_folder = r"C:\Users\NIC\OneDrive\Desktop\Election\Target\236"
convert_excel_to_pdf_with_borders(source_folder, target_folder)