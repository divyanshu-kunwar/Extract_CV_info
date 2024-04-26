import os
import re
from openpyxl import Workbook
import PyPDF2
from docx import Document
import comtypes.client
import win32com.client

def extract_text_from_doc(doc_file):
    """Extracts text from a DOC file using pywin32."""
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(doc_file)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
        return text
    except Exception as e:
        print(f"Error processing DOC {doc_file}: {e}")
        return ""

def process_cvs(cv_folder, output_filename):
    """Processes CVs in a folder and saves the extracted information in .xlsx format."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Email", "Phone Number", "Text"])  # Header row

    for root, _, files in os.walk(cv_folder):
        for filename in files:
            if filename.endswith(".pdf"):
                text = extract_text_from_pdf(os.path.join(root, filename))
            elif filename.endswith(".docx"):
                text = extract_text_from_docx(os.path.join(root, filename))
            elif filename.endswith(".doc"):
                text = extract_text_from_doc(os.path.join(root, filename))
            else:
                print(f"Skipping unsupported file: {filename}")
                continue

            info = extract_info(text)
            ws.append([info["email"], info["phone_number"], info["text"]])

    wb.save(output_filename)  # Save the Excel file

if __name__ == "__main__":
    cv_folder = r"D:\Sample2_2024\Sample2"  # Path to CV folder
    output_filename = "extracted_info5.xlsx"
    process_cvs(cv_folder, output_filename)
