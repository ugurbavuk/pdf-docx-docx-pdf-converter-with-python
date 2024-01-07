import os
from tkinter import Tk, filedialog, Label, Button
import fitz  # PyMuPDF
from docx import Document

class PDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to DOCX Converter")

        self.label = Label(root, text="PDF to DOCX Converter", font=("Helvetica", 16))
        self.label.pack(pady=10)

        self.convert_button = Button(root, text="Convert PDF to DOCX", command=self.convert_pdf_to_docx)
        self.convert_button.pack(pady=20)

        self.reverse_button = Button(root, text="Convert DOCX to PDF", command=self.convert_docx_to_pdf)
        self.reverse_button.pack(pady=20)

    def convert_pdf_to_docx(self):
        file_path = filedialog.askopenfilename(title="Select a PDF file", filetypes=[("PDF Files", "*.pdf")])

        if file_path:
            doc = Document()
            pdf_document = fitz.open(file_path)

            for page_number in range(pdf_document.page_count):
                page = pdf_document.load_page(page_number)
                text = page.get_text()
                doc.add_paragraph(text)

            docx_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
            if docx_path:
                doc.save(docx_path)
                print("Conversion successful.")

    def convert_docx_to_pdf(self):
        file_path = filedialog.askopenfilename(title="Select a DOCX file", filetypes=[("Word Files", "*.docx")])

        if file_path:
            doc = Document(file_path)
            pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])

            if pdf_path:
                pdf_document = fitz.open()
                for paragraph in doc.paragraphs:
                    pdf_page = pdf_document.new_page()
                    pdf_page.insert_text((10, 10), paragraph.text)

                pdf_document.save(pdf_path)
                print("Conversion successful.")


if __name__ == "__main__":
    root = Tk()
    converter = PDFConverter(root)
    root.mainloop()
