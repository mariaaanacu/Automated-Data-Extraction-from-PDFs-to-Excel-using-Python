import re
import os
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog, messagebox
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator

def extract_information(pdf_path):
    resource_manager = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(resource_manager, laparams=laparams)
    interpreter = PDFPageInterpreter(resource_manager, device)

    product_details = []
    seen_codes = set()

    with open(pdf_path, 'rb') as pdf_file:
        for page in PDFPage.get_pages(pdf_file, check_extractable=True):
            interpreter.process_page(page)
            layout = device.get_result()

            for lt_obj in layout:
                if isinstance(lt_obj, LTTextBox):
                    text = lt_obj.get_text()

                    # Pattern to find product details and code
                    pattern = re.compile(r'(.*?)\s*COD\. (\d{6})(?=\s*•|\s*$)', re.DOTALL)
                    matches = pattern.findall(text)

                    for match in matches:
                        details = match[0].strip()
                        details = re.sub(r'\s+', ' ', details)  # Replace multiple spaces/newlines with a single space
                        code = match[1].strip()

                        if code not in seen_codes:
                            seen_codes.add(code)
                            # Split characteristics using bullet points (•)
                            characteristics = [c.strip() for c in details.split('•') if c.strip()]
                            product_details.append((characteristics, code))

    return product_details

def create_excel(product_details, excel_path):
    try:
        wb = Workbook()
        ws = wb.active
        ws.append(["Cod Produs", "Name", "Caracteristica 1", "Caracteristica 2", "Caracteristica 3", "Caracteristica 4", "Caracteristica 5", "Caracteristica 6", "Caracteristica 7", "Caracteristica 8", "Caracteristica 9"])

        for details in product_details:
            row = [
                details[1]   # Cod Produs
            ]
            row.extend(details[0])  # Add all characteristics

            # Complete the rest of the columns with empty strings if there are fewer than 10 characteristics
            if len(details[0]) < 10:
                row.extend([''] * (10 - len(details[0])))

            ws.append(row)

        wb.save(excel_path)
        messagebox.showinfo("Succes", "Fișierul Excel a fost creat cu succes!")
    except Exception as e:
        messagebox.showerror("Eroare", f"A apărut o eroare la crearea fișierului Excel: {e}")

def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        product_details = extract_information(pdf_path)
        # Get the directory of the PDF file
        dir_path = os.path.dirname(pdf_path)
        # Create the full path for the Excel file in the same directory
        excel_path = os.path.join(dir_path, "output.xlsx")
        create_excel(product_details, excel_path)

# Crearea interfeței grafice
root = tk.Tk()
root.title("PDF to Excel Converter")

frame = tk.Frame(root)
frame.pack(padx=20, pady=20)

browse_button = tk.Button(frame, text="Selectează PDF", command=browse_pdf)
browse_button.pack()

root.mainloop()
