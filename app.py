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
