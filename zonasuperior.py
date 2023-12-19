import os
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd

def pdf_to_image(pdf_path, page_number, region):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number - 1)
    mat = fitz.Matrix(2, 2)  # Adjust the resolution as needed

    # Render the page as a pixmap
    pix = page.get_pixmap(matrix=mat)

    # Convert the pixmap to a PIL image
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Crop the region of interest from the image
    left, top, right, bottom = region
    img = img.crop((left, top, right, bottom))

    return img

def ocr_image_to_text(image, lang='eng'):
    # Perform OCR on the image using pytesseract
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    text = pytesseract.image_to_string(image, lang=lang)

    return text

def process_all_pdfs(directory, keyword1, keyword2, excel_file):
    # Create an empty DataFrame to store the results
    results_df = pd.DataFrame(columns=['Filename', 'Page', 'Text'])

    # Iterate through all files in the specified directory
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)

            # Get the total number of pages in the PDF
            with fitz.open(pdf_path) as doc:
                num_pages = doc.page_count

            # Process each page in the PDF
            for page_number in range(1, num_pages + 1):
                # Define the region for cropping (adjust coordinates as needed)
                region = (0, 0, 1700, 1700)

                # Convert PDF to image
                image = pdf_to_image(pdf_path, page_number, region)

                # Save the cropped image as a PNG file
                output_path = os.path.join(directory, f"{filename}_page_{page_number}_cropped.png")
                image.save(output_path)

                # Perform OCR on the image
                text = ocr_image_to_text(image)

                # Check if both keywords are present in the text
                if keyword1 in text and keyword2 in text:
                    # Append the result to the DataFrame
                    results_df = results_df._append({'Filename': filename, 'Page': page_number, 'Text': text}, ignore_index=True)

    # Save the results to the Excel file
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
        results_df.to_excel(writer, sheet_name='Sheet2', index=False, header=False)

# Specify the directory containing the PDF files
pdf_directory = r"C:\Users\anton\Documents\OCR\ubica"

# Specify the keywords to check for in the text
keyword1 = "Ejercicio"
keyword2 = "Periodo"

# Specify the Excel file
excel_file = "resultados.xlsx"

# Call the function to process all PDFs in the directory and append results to the Excel file
process_all_pdfs(pdf_directory, keyword1, keyword2, excel_file)
print(text)