import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd
import os

def pdf_to_image(pdf_path, page_number, region):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number - 1)
    mat = fitz.Matrix(2, 2)  # Ajusta la resolución según sea necesario

    # Renderiza la página como una imagen (pixmap)
    pix = page.get_pixmap(matrix=mat)

    # Convierte el pixmap a una imagen PIL (Python Imaging Library)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Recorta la región de interés de la imagen
    left, top, right, bottom = region
    img = img.crop((left, top, right, bottom))

    return img

def ocr_image_to_text(image, lang='eng'):
    # Realiza OCR en la imagen utilizando pytesseract
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    text = pytesseract.image_to_string(image, lang=lang)

    return text

def find_and_store_text(pdf_path, search_text, results):
    pdf_doc = fitz.open(pdf_path)

    for page_number in range(1, 1 + pdf_doc.page_count):
        # Convertir la región seleccionada del PDF a una imagen
        image = pdf_to_image(pdf_path, page_number, (0, 0, 1600, 1000))

        # Realizar OCR en la imagen y obtener el texto
        result_text = ocr_image_to_text(image)

        # Buscar la cadena específica en el texto
        index = result_text.find(search_text)

        if index != -1:
            # Encontrar el inicio y el final de la línea
            start_index = index
            end_index = result_text.find('\n', start_index)

            # Si no hay un próximo salto de línea, tomar hasta el final del texto
            if end_index == -1:
                end_index = len(result_text)

            # Eliminar los puntos suspensivos detrás de "ejercidas" y agregar ":"
            relevant_text = result_text[start_index:end_index].replace("ejercidas...", "ejercidas:") 

            # Agregar ":" al principio de la cadena desde el final del texto hasta el primer espacio
            first_space_index = relevant_text.rfind(' ')
            relevant_text = ':' + relevant_text[first_space_index + 1:]

            # Agregar el resultado a la lista de resultados
            results.append((pdf_path, page_number, relevant_text))

if __name__ == "__main__":
    # Carpeta que contiene los archivos PDF
    folder_path = r"C:\Users\anton\Documents\OCR\ubica"

    # Texto a buscar
    search_text = "Ingresos computables correspondientes al conjunto de las actividades ejercidas"

    # Lista para almacenar los resultados
    all_results = []

    # Iterar sobre los archivos en la carpeta
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            find_and_store_text(pdf_path, search_text, all_results)

    # Crear un DataFrame con los resultados
    df = pd.DataFrame(all_results, columns=['Archivo PDF', 'Número de Página', 'Texto Encontrado'])

    # Exportar el DataFrame a un archivo Excel
    output_excel = "resultados.xlsx"
    df.to_excel(output_excel, index=False)
    print(f"Resultados exportados a {output_excel}")

