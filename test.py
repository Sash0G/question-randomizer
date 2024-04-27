import pdf2image
from PIL import Image
import pytesseract
from docx import Document

def pdf_to_img(pdf_file):
    return pdf2image.convert_from_path(pdf_file)


def ocr_core(file):
    text = pytesseract.image_to_string(file, lang='bul')
    return text


def print_pages(pdf_file):
    images = pdf_to_img(pdf_file)
    for pg, img in enumerate(images):
        text = ocr_core(img)
        print(text)
        try: doc.add_paragraph(text)
        except: continue

doc = Document()
print_pages('/media/sasho_g/SSD/Users/GRIGS/Downloads/МУ Варна - Биология част 1 - 2022.pdf')
doc.save('output.docx')