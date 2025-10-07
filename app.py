from flask import Flask, request, render_template
import os, re
from docx import Document
import pdfplumber
from pptx import Presentation
from openpyxl import load_workbook

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

PORT = int(os.environ.get('PORT', 10000))

# --- Actual helper functions ---
def search_in_docx(file_path, query):
    results = []
    try:
        doc = Document(file_path)
        for para in doc.paragraphs:
            if query.lower() in para.text.lower():
                results.append(para.text.strip())
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return results

def search_in_pdf(file_path, query):
    results = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text and query.lower() in text.lower():
                    match = re.search(f".{{0,30}}{re.escape(query)}.{{0,30}}", text, re.IGNORECASE)
                    if match:
                        results.append(f"[Page {page_num}] ...{match.group(0)}...")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return results

def search_in_pptx(file_path, query):
    results = []
    try:
        prs = Presentation(file_path)
        for i, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                if hasattr(shape, "text") and query.lower() in shape.text.lower():
                    snippet = re.search(f".{{0,30}}{re.escape(query)}.{{0,30}}", shape.text, re.IGNORECASE)
                    if snippet:
                        results.append(f"[Slide {i}] ...{snippet.group(0)}...")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return results

def search_in_excel(file_path, query):
    results = []
    try:
        wb = load_workbook(file_path, data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if isinstance(cell, str) and query.lower() in cell.lower():
                        results.append(f"[{sheet}] {cell.strip()}")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return results

# --- Routes ---
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/search', methods=['GET', 'POST'])
def search():
    query = request.form.get('query', '').strip() if request.method == 'POST' else ''
    results = []

    if query:
        print("Searching for:", query)
        print("UPLOAD_FOLDER:", UPLOAD_FOLDER)
        for root, _, files in os.walk(UPLOAD_FOLDER):
            print("Files found:", files)
            for file in files:
                file_path = os.path.join(root, file)
                ext = os.path.splitext(file)[1].lower()
                file_results = []

                if ext == ".docx":
                    file_results = search_in_docx(file_path, query)
                elif ext == ".pdf":
                    file_results = search_in_pdf(file_path, query)
                elif ext == ".pptx":
                    file_results = search_in_pptx(file_path, query)
                elif ext in [".xlsx", ".xls"]:
                    file_results = search_in_excel(file_path, query)

                if file_results:
                    results.append({"file": file, "matches": file_results})

    return render_template("index.html", query=query, results=results)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=PORT, debug=False)
