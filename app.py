from flask import Flask, request, render_template_string, send_from_directory
from docxtpl import DocxTemplate
import os

app = Flask(__name__)

UPLOAD_FOLDER = "outputs"
TEMPLATE_FILE = "holistic template.docx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Makale Oluştur</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 900px;
            margin: 30px auto;
            line-height: 1.5;
        }
        input, textarea {
            width: 100%;
            padding: 8px;
            margin-top: 4px;
            margin-bottom: 16px;
            box-sizing: border-box;
        }
        textarea {
            min-height: 120px;
        }
        button {
            padding: 10px 18px;
            font-size: 16px;
        }
        h1 {
            margin-bottom: 24px;
        }
    </style>
</head>
<body>
    <h1>Makale Oluştur</h1>

    <form method="POST">
        <label>Makale Başlığı</label>
        <input type="text" name="title" required>

        <label>Yazar Bloğu</label>
        <textarea name="author_block" placeholder="Kör hakemlik için boş bırakabilir veya Anonymous Author(s) yazdırabiliriz."></textarea>

        <label>Abstract</label>
        <textarea name="abstract" required></textarea>

        <label>Keywords</label>
        <input type="text" name="keywords" required>

        <label>Body</label>
        <textarea name="body" required placeholder="Yazar ana metni buraya yapıştıracak. Başlıklar da bunun içinde olabilir."></textarea>

        <label>Acknowledgements</label>
        <textarea name="acknowledgements"></textarea>

        <label>Funding</label>
        <textarea name="funding"></textarea>

        <label>Conflict of Interest</label>
        <textarea name="conflict_of_interest"></textarea>

        <label>References</label>
        <textarea name="references" required></textarea>

        <label>
            <input type="checkbox" name="blind_review" value="yes">
            Kör hakem sürümü üret
        </label>

        <br><br>
        <button type="submit">Word Oluştur</button>
    </form>
</body>
</html>
"""

RESULT_HTML = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Dosya Hazır</title>
</head>
<body style="font-family: Arial, sans-serif; max-width: 900px; margin: 30px auto;">
    <h2>Word dosyası hazır</h2>
    <p><a href="/download/{{ filename }}">Dosyayı indir</a></p>
    <p><a href="/">Yeni makale oluştur</a></p>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        author_block = request.form.get("author_block", "").strip()
        abstract = request.form.get("abstract", "").strip()
        keywords = request.form.get("keywords", "").strip()
        body = request.form.get("body", "").strip()
        acknowledgements = request.form.get("acknowledgements", "").strip()
        funding = request.form.get("funding", "").strip()
        conflict_of_interest = request.form.get("conflict_of_interest", "").strip()
        references = request.form.get("references", "").strip()
        blind_review = request.form.get("blind_review")

        if blind_review == "yes":
            author_block = "Anonymous Author(s)"

        doc = DocxTemplate(TEMPLATE_FILE)

        context = {
            "title": title,
            "author_block": author_block,
            "abstract": abstract,
            "keywords": keywords,
            "body": body,
            "acknowledgements": acknowledgements,
            "funding": funding,
            "conflict_of_interest": conflict_of_interest,
            "references": references,
        }

        doc.render(context)

        filename = "generated_article.docx"
        output_path = os.path.join(UPLOAD_FOLDER, filename)
        doc.save(output_path)

        return render_template_string(RESULT_HTML, filename=filename)

    return render_template_string(HTML_FORM)

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
