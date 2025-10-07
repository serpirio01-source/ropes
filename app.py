from flask import Flask, request, render_template
import os, re
from docx import Document
import pdfplumber
from pptx import Presentation
from openpyxl import load_workbook

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

PORT = int(os.environ.get('PORT', 10000))

# --- All helper functions FIRST ---
def search_in_docx(file_path, query): ...
def search_in_pdf(file_path, query): ...
def search_in_pptx(file_path, query): ...
def search_in_excel(file_path, query): ...

# --- Then your routes ---
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/search', methods=['GET', 'POST'])
def search():
    query = request.form.get('query', '').strip() if request.method == 'POST' else ''
    results = []

    if query:
        for root, _, files in os.walk(UPLOAD_FOLDER):
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
