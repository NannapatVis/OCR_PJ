import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
import pytesseract as tess
from pathlib import Path
import re
import time
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph

# Set the path to Tesseract OCR
tess.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Keywords to search and highlight
keywords = ['John Smith']

# Start timer for processing time
start_time = time.time()

# Load PDF for text extraction and page search
pdf_path = 'example.pdf'
pdf = PdfReader(pdf_path)

# Create a folder to store results
output_folder = 'output_files'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create a folder to store temporary images for OCR
ocr_image_folder = os.path.join(output_folder, 'ocr_images')
if not os.path.exists(ocr_image_folder):
    os.makedirs(ocr_image_folder)

# Search for pages that contain any of the keywords (case-insensitive)
keyword_pages = []
for page_num, page in enumerate(pdf.pages):
    page_text = page.extract_text()
    
    # If page text extraction fails, fallback to OCR
    if not page_text:
        images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, output_folder=ocr_image_folder)
        if images:
            page_text = tess.image_to_string(images[0], lang='eng')
    
    if page_text:
        for keyword in keywords:
            if re.search(rf'\b{keyword}\b', page_text, re.IGNORECASE):
                keyword_pages.append(page_num)
                break  # Stop once a keyword is found in the page

# Save pages that contain any of the keywords to a new PDF file
pdf_writer = PdfWriter()
for page_num in keyword_pages:
    page_object = PdfReader(pdf_path).pages[page_num]
    pdf_writer.add_page(page_object)

with Path(os.path.join(output_folder, 'keyword_pages.pdf')).open(mode='wb') as output_file_2:
    pdf_writer.write(output_file_2)

# Use PyMuPDF to highlight the keywords in the new PDF
pdf_document = fitz.open(pdf_path)
highlighted_pdf_path = os.path.join(output_folder, 'highlighted_keyword_pages.pdf')
highlighted_pdf_document = fitz.open()

for page_num in keyword_pages:
    # Load the page and create a new page in the highlighted PDF
    page = pdf_document.load_page(page_num)
    highlighted_page = highlighted_pdf_document.new_page(width=page.rect.width, height=page.rect.height)
    
    # Copy the content from the original page
    highlighted_page.show_pdf_page(highlighted_page.rect, pdf_document, page_num)
    
    # Highlight the instances of the keywords (case-insensitive)
    page_text = page.get_text()
    if not page_text:  # If no text is extracted, use OCR for highlighting
        images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, output_folder=ocr_image_folder)
        if images:
            page_text = tess.image_to_string(images[0], lang='eng')
    
    for keyword in keywords:
        search_term = keyword.lower()
        start = 0
        while True:
            start = page_text.lower().find(search_term, start)
            if start == -1:
                break
            end = start + len(search_term)
            rects = page.search_for(page_text[start:end])
            for rect in rects:
                highlighted_page.add_highlight_annot(rect)
            start = end

# Save the highlighted PDF with the keywords highlighted
highlighted_pdf_document.save(highlighted_pdf_path)
highlighted_pdf_document.close()
pdf_document.close()

# Capture processing time
end_time = time.time()
processing_time = end_time - start_time
print(f"Processing time: {processing_time:.2f} seconds")

# Extract sentences that contain the keywords and include page numbers
pages_sentences = []
for page_num, page in enumerate(pdf.pages):
    page_text = page.extract_text()
    
    # Use OCR if text extraction fails
    if not page_text:
        images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, output_folder=ocr_image_folder)
        if images:
            page_text = tess.image_to_string(images[0], lang='eng')

    if page_text:
        for keyword in keywords:
            # Find and extract sentences with the keywords
            sentences = re.split(r'(?<=[.!?])\s+', page_text)
            filtered_sentences = [sentence.strip() for sentence in sentences if re.search(rf'\b{keyword}\b', sentence, re.IGNORECASE)]
            
            if filtered_sentences:
                formatted_sentence = f'Page {page_num + 1}: ' + ' '.join(filtered_sentences)
                pages_sentences.append(formatted_sentence)

# Save the sentences containing the keywords to a text file
text = '\n'.join(pages_sentences)
with Path(os.path.join(output_folder, 'keyword_sentences.txt')).open(mode='w', encoding='utf-8') as output_file_3:
    output_file_3.write(text)

# Convert keyword sentences to PDF without highlights (normal text)
pdf_file = os.path.join(output_folder, 'keyword_sentences.pdf')
doc = SimpleDocTemplate(pdf_file, pagesize=letter)
styles = getSampleStyleSheet()
story = []

# Create PDF content from the keyword sentences
for line in pages_sentences:
    story.append(Paragraph(line, styles['BodyText']))

# Build the PDF document
doc.build(story)

# Clean up the OCR image folder after use
for file in os.listdir(ocr_image_folder):
    file_path = os.path.join(ocr_image_folder, file)
    if os.path.isfile(file_path):
        os.remove(file_path)
os.rmdir(ocr_image_folder)

print("Keyword sentences saved to PDF.")
