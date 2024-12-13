from parser import extract_data_from_pdf, save_to_excel, save_to_db, extract_images_from_pdf
import os

pdf_file = "file2.pdf"
xlsx_file = "result.xlsx"
db_file = "result.db"
images_folder = "images"

if __name__ == "__main__":
    os.makedirs(images_folder, exist_ok=True)
    extracted_data = extract_data_from_pdf(pdf_file)

    save_to_excel(extracted_data, xlsx_file)
    save_to_db(extracted_data, db_file)

    extract_images_from_pdf(pdf_file, images_folder)
