import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract as tess
from pathlib import Path
import re
import time
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.colors import red
import pdfplumber  

# Set the path to Tesseract OCR
tess.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Default keywords (can be left empty if you want to force input from Terminal)
keywords = ['Mary Jane']  

# Check if keywords are empty or contain only empty strings
if not keywords or all(kw.strip() == '' for kw in keywords):
    # Prompt the user for keywords if no valid keywords are provided in the code
    user_input = input("Enter keywords separated by commas (required): ")
    
    # Continue prompting until valid input is provided
    while not user_input.strip():
        user_input = input("No keywords entered. Please enter at least one keyword: ")
    
    keywords = [kw.strip() for kw in user_input.split(',')]

# Start timer for processing time
start_time = time.time()

# Load PDF for text extraction and page search
pdf_path = r'Ex\3.-The-Adventures-of-Sherlock-Holmes-Author-Arthur-Conan-Doyle.pdf'

# Create a folder to store results
output_folder = 'result OCR'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create a folder to store temporary images for OCR
ocr_image_folder = os.path.join(output_folder, 'ocr_images')
if not os.path.exists(ocr_image_folder):
    os.makedirs(ocr_image_folder)

# Search for keywords and store relevant paragraphs
keyword_counts = {}  # Dictionary to store keyword occurrences by page
pages_sentences = []
pages_sentences_txt = []  # This list will store sentences with ** highlighting for the .txt file

with pdfplumber.open(pdf_path) as pdf:
    for page_num, page in enumerate(pdf.pages):
        page_text = page.extract_text()

        # If page text extraction fails, fallback to OCR
        if not page_text:
            images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, output_folder=ocr_image_folder)
            if images:
                page_text = tess.image_to_string(images[0], lang='tha+eng')

        if page_text:
            # Split text into paragraphs
            paragraphs = page_text.split('\n\n')  # Adjust to your PDF's formatting
            for para_num, paragraph in enumerate(paragraphs):
                for keyword in keywords:
                    if re.search(rf'\b{re.escape(keyword)}\b', paragraph, re.IGNORECASE):
                        if page_num not in keyword_counts:
                            keyword_counts[page_num] = 0
                        keyword_counts[page_num] += len(re.findall(rf'\b{re.escape(keyword)}\b', paragraph, re.IGNORECASE))
                        formatted_sentence = f'Page {page_num + 1}: {paragraph.strip()}'
                        pages_sentences.append(formatted_sentence)

                        # For the .txt file, add ** around the keywords
                        highlighted_sentence = formatted_sentence
                        for keyword in keywords:
                            highlighted_sentence = re.sub(rf'\b{re.escape(keyword)}\b', f'**{keyword}**', highlighted_sentence, flags=re.IGNORECASE)
                        pages_sentences_txt.append(highlighted_sentence)

# Save the sentences containing the keywords to a text file with ** highlighting
text = '\n'.join(pages_sentences_txt)
with Path(os.path.join(output_folder, 'keyword_sentences.txt')).open(mode='w', encoding='utf-8') as output_file_3:
    output_file_3.write(text)

# Register a Thai font
pdfmetrics.registerFont(TTFont('THSarabunNew', r'C:\Users\aaa\Desktop\OCR\Font\THSarabunNew.ttf'))

# Create a stylesheet and assign the Thai font to the body text
styles = getSampleStyleSheet()
styles['BodyText'].fontName = 'THSarabunNew'
styles['BodyText'].fontSize = 12

# Create a PDF file
pdf_file = os.path.join(output_folder, 'keyword_sentences.pdf')
doc = SimpleDocTemplate(pdf_file, pagesize=letter)
story = []

# Add keyword summary to the PDF document
story.append(Paragraph(f"Summary of Keywords: {', '.join(keywords)}", styles['Heading2']))  # เพิ่มการแสดง keyword
story.append(Spacer(1, 12))

for page_num, count in keyword_counts.items():
    summary_line = f"Page {page_num + 1}: {count} occurrence(s)"
    story.append(Paragraph(summary_line, styles['BodyText']))
    story.append(Spacer(1, 12))

# Add a spacer before adding the actual keyword sentences
story.append(Spacer(1, 24))

# Highlight keywords in red when creating PDF content from the keyword sentences
for line in pages_sentences:
    # Remove both "Paragraph" and "Page" info, but keep the page number
    clean_line = re.sub(r'Paragraph \d+: ', '', line)  # Remove paragraph info
    clean_line = re.sub(r'Page (\d+), ', r'Page \1: ', clean_line)  # Format page number
    
    for keyword in keywords:
        # Case-insensitive search but retain the original case for output
        pattern = re.compile(rf'({re.escape(keyword)})', re.IGNORECASE)
        # Use <b> and <font color='red'> instead of <font color="red"><b>
        clean_line = pattern.sub(r'<b><font color="red">\1</font></b>', clean_line)
    
    # Append the clean sentence to the story for the PDF
    story.append(Paragraph(clean_line, styles['BodyText']))

# Build the PDF document
doc.build(story)

# Clean up the OCR image folder after use
for file in os.listdir(ocr_image_folder):
    file_path = os.path.join(ocr_image_folder, file)
    if os.path.isfile(file_path):
        os.remove(file_path)
os.rmdir(ocr_image_folder)

# Calculate and print the processing time
end_time = time.time()
processing_time = end_time - start_time
print(f"Processing completed in {processing_time:.2f} seconds.")

print("Keyword sentences saved to PDF and text files.")