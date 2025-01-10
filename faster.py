import camelot
import fitz 
import pytesseract
from PIL import Image
from bs4 import BeautifulSoup
import pypandoc
import os
import logging
import layoutparser
from layoutparser import Detectron2LayoutModel

layout_model = Detectron2LayoutModel(
    config_path="lp://PubLayNet/faster_rcnn_R_50_FPN_3x/config",
    label_map={0: "Text", 1: "Title", 2: "List", 3: "Table", 4: "Figure"},
    extra_config=["MODEL.DEVICE", "cuda" if torch.cuda.is_available() else "cpu"]
)


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def extract_tables(pdf_path, output_dir):
    logging.info("Starting table extraction...")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    tables = []
    try:
        tables = camelot.read_pdf(pdf_path, flavor='lattice', pages='all')
        if len(tables) == 0:
            logging.warning("No structured tables found, attempting unstructured extraction...")
            tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

        for i, table in enumerate(tables):
            table.df.to_csv(os.path.join(output_dir, f'table_{i}.csv'), index=False)
        logging.info(f"Extracted {len(tables)} tables.")
    except Exception as e:
        logging.error(f"Failed to extract tables: {e}")

    return tables


def extract_content_with_layout(pdf_path, output_dir, layout_model):

    logging.info("Starting content extraction with layout analysis...")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            img_path = os.path.join(output_dir, f'page_{page_num}.png')
            pix.save(img_path)

            img = cv2.imread(img_path)
            layout = layout_model.detect(img)
            structured_text = ""

            for block in layout:
                if block['type'] == 'Text':
                    bbox = block['bbox']
                    cropped_img = img[int(bbox[1]):int(bbox[3]), int(bbox[0]):int(bbox[2])]
                    text = pytesseract.image_to_string(cropped_img, lang='rus+eng')
                    structured_text += f"{text}\n\n"

            with open(os.path.join(output_dir, f'page_{page_num}_structured.txt'), 'w', encoding='utf-8') as f:
                f.write(structured_text)

        except Exception as e:
            logging.error(f"Failed content extraction on page {page_num}: {e}")

def layout_analysis(pdf_path, output_dir):

    logging.info("Starting layout analysis...")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            img_path = os.path.join(output_dir, f'page_{page_num}.png')
            pix.save(img_path)

            hocr_output = pytesseract.image_to_pdf_or_hocr(img_path, extension='hocr')
            with open(os.path.join(output_dir, f'page_{page_num}.hocr'), 'wb') as f:
                f.write(hocr_output)
        except Exception as e:
            logging.error(f"Failed layout analysis for page {page_num}: {e}")
            
def add_tables_to_doc(doc, table_bbox, image_path):
    img = cv2.imread(image_path)
    contours, _ = cv2.findContours(table_bbox, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        table_region = img[y:y+h, x:x+w]
        table_text = pytesseract.image_to_string(table_region, config='--psm 6')
        rows = table_text.split("\n")
        table = doc.add_table(rows=len(rows), cols=len(rows[0].split("\t")))
        table.style = 'Table Grid'

        for i, row in enumerate(rows):
            for j, cell in enumerate(row.split("\t")):
                table.cell(i, j).text = cell



def generate_html(content_dir, tables_dir, layout_dir, output_html):

    logging.info("Generating structured HTML...")
    try:
        html = "<html><head><title>PDF Content</title><style>"
        html += "body { font-family: Arial, sans-serif; }"
        html += ".page { margin-bottom: 20px; }"
        html += ".table { margin: 10px 0; border: 1px solid black; }"
        html += "</style></head><body>"

        for txt_file in sorted(os.listdir(content_dir)):
            if txt_file.endswith("_structured.txt"):
                with open(os.path.join(content_dir, txt_file), 'r', encoding='utf-8') as f:
                    html += f"<div class='page'>{f.read()}</div>"

        for table_file in sorted(os.listdir(tables_dir)):
            if table_file.endswith(".csv"):
                html += f"<div class='table'><h3>{table_file}</h3>"
                html += "<table border='1'>"
                with open(os.path.join(tables_dir, table_file), 'r', encoding='utf-8') as f:
                    rows = f.readlines()
                    for row in rows:
                        html += "<tr>" + "".join([f"<td>{cell}</td>" for cell in row.split(',')]) + "</tr>"
                html += "</table></div>"

        html += "</body></html>"

        with open(output_html, 'w', encoding='utf-8') as f:
            f.write(html)
        logging.info(f"Structured HTML saved at {output_html}")

    except Exception as e:
        logging.error(f"Failed to generate structured HTML: {e}")



def html_to_word(html_path, word_path):
    logging.info("Converting HTML to Word...")
    if not os.path.exists(html_path):
        logging.error(f"HTML file does not exist: {html_path}")
        return
    try:
        pypandoc.convert_file(html_path, 'docx', outputfile=word_path)
        logging.info(f"Word document saved at {word_path}")
    except Exception as e:
        logging.error(f"Failed to convert HTML to Word: {e}")


def main(pdf_path, output_dir):

    try:
        tables_dir = os.path.join(output_dir, 'tables')
        content_dir = os.path.join(output_dir, 'content')
        layout_dir = os.path.join(output_dir, 'layout')
        extract_tables(pdf_path, tables_dir)       
        extract_content_with_layout(pdf_path, content_dir, layout_model)
        html_path = os.path.join(output_dir, 'output.html')
        generate_html(content_dir, tables_dir, layout_dir, html_path)
        
        word_path = os.path.join(output_dir, 'output.docx')
        html_to_word(html_path, word_path)
    
    except Exception as e:
        logging.error(f"Failed to process PDF: {e}")

if __name__ == "__main__":
    pdf_path = r'docs\sample test 1.pdf'  # Replace with your PDF file path
    output_dir = 'output'
    main(pdf_path, output_dir)
