import pdfplumber
import fitz  # PyMuPDF
import sqlite3
from openpyxl import Workbook
import os

def extract_data_from_pdf(pdf_path):
    extracted_data = {"text": [], "tables": []}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    for line in text.splitlines():
                        extracted_data["text"].append((page_number, line))
                
                tables = page.extract_tables()
                for table in tables:
                    extracted_data["tables"].append({"page_number": page_number, "table": table})
        return extracted_data
    except BaseException:
        print(f"File '{pdf_path}' does not exist")
    #return extracted_data

def extract_images_from_pdf(pdf_path, output_folder):
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc, start=1):
        for img_index, img in enumerate(page.get_images(full=True), start=1):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            ext = base_image["ext"]
            image_filename = os.path.join(output_folder, f"page_{i}_image_{img_index}.{ext}")
            with open(image_filename, "wb") as img_file:
                img_file.write(image_bytes)
    print(f"Images saved in folder '{output_folder}'")

def save_to_excel(data, xlsx_path):
    workbook = Workbook()

    sheet_text = workbook.active
    sheet_text.title = "Text"
    sheet_text.append(["Page Number", "Line Text"])
    for page_number, line in data["text"]:
        sheet_text.append([page_number, line])

    sheet_tables = workbook.create_sheet(title="Tables")
    sheet_tables.append(["Page Number", "Row Data"])
    for table_data in data["tables"]:
        page_number = table_data["page_number"]
        table = table_data["table"]
        for row in table:
            sheet_tables.append([page_number, ", ".join(row)])

    workbook.save(xlsx_path)
    print(f"Data saved to '{xlsx_path}'")

def save_to_db(data, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pdf_text (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            page_number INTEGER,
            line_text TEXT
        )
    """)
    cursor.executemany("INSERT INTO pdf_text (page_number, line_text) VALUES (?, ?)", data["text"])

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pdf_tables (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            page_number INTEGER,
            row_data TEXT
        )
    """)
    for table_data in data["tables"]:
        page_number = table_data["page_number"]
        table = table_data["table"]
        for row in table:
            cursor.execute("INSERT INTO pdf_tables (page_number, row_data) VALUES (?, ?)", (page_number, ", ".join(row)))

    conn.commit()
    conn.close()
    print(f"Data saved to database '{db_path}'")
