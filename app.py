from flask import Flask, request, render_template_string, send_from_directory
from docxtpl import DocxTemplate
from docx import Document
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
TEMPLATE_FILE = "holistic template.docx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# -------------------------
# APA FORMAT
# -------------------------
def format_apa(refs):
    formatted = []

    for r in refs:
        if r["type"] == "article":
            apa = f'{r["author"]} ({r["year"]}). {r["title"]}. {r["journal"]}, {r["volume"]}({r["issue"]}), {r["pages"]}.'
        elif r["type"] == "book":
            apa = f'{r["author"]} ({r["year"]}). {r["title"]}. {r["publisher"]}.'
        else:
            apa = r["title"]

        formatted.append(apa)

    return "\n".join(formatted)

# -------------------------
# WORD BODY OKUMA
# -------------------------
def read_docx(file_path):
    doc = Document(file_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    return "\n".join(text)

# -------------------------
# FORM HTML
# -------------------------
HTML = """
<h2>Makale Sistemi</h2>

<form method="POST" enctype="multipart/form-data">

Başlık:<br>
<input name="title"><br><br>

Abstract:<br>
<textarea name="abstract"></textarea><br><br>

Keywords:<br>
<input name="keywords"><br><br>

Body (Word yükle):<br>
<input type="file" name="body_file"><br><br>

Acknowledgements:<br>
<textarea name="ack"></textarea><br><br>

Funding:<br>
<textarea name="funding"></textarea><br><br>

Conflict:<br>
<textarea name="conflict"></textarea><br><br>

<h3>Referans (1 adet örnek)</h3>

Tür:
<select name="ref_type">
<option value="article">Makale</option>
<option value="book">Kitap</option>
</select><br>

Yazar:<input name="ref_author"><br>
Yıl:<input name="ref_year"><br>
Başlık:<input name="ref_title"><br>
Dergi:<input name="ref_journal"><br>
Cilt:<input name="ref_volume"><br>
Sayı:<input name="ref_issue"><br>
Sayfa:<input name="ref_pages"><br>
Yayıncı:<input name="ref_publisher"><br><br>

<input type="checkbox" name="blind"> Kör hakem<br><br>

<button>Oluştur</button>

</form>
"""

# -------------------------
# ROUTE
# -------------------------
@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        title = request.form.get("title")
        abstract = request.form.get("abstract")
        keywords = request.form.get("keywords")
        ack = request.form.get("ack")
        funding = request.form.get("funding")
        conflict = request.form.get("conflict")

        # Kör hakem
        author_block = "Anonymous Author(s)" if request.form.get("blind") else ""

        # -------------------------
        # WORD BODY
        # -------------------------
        file = request.files["body_file"]
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        body = read_docx(filepath)

        # -------------------------
        # REFERENCES (STRUCTURED)
        # -------------------------
        ref = {
            "type": request.form.get("ref_type"),
            "author": request.form.get("ref_author"),
            "year": request.form.get("ref_year"),
            "title": request.form.get("ref_title"),
            "journal": request.form.get("ref_journal"),
            "volume": request.form.get("ref_volume"),
            "issue": request.form.get("ref_issue"),
            "pages": request.form.get("ref_pages"),
            "publisher": request.form.get("ref_publisher"),
        }

        references = format_apa([ref])

        # -------------------------
        # TEMPLATE
        # -------------------------
        doc = DocxTemplate(TEMPLATE_FILE)

        context = {
            "title": title,
            "author_block": author_block,
            "abstract": abstract,
            "keywords": keywords,
            "body": body,
            "acknowledgements": ack,
            "funding": funding,
            "conflict_of_interest": conflict,
            "references": references
        }

        doc.render(context)

        output_file = "final.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_file)
        doc.save(output_path)

        return f'<a href="/download/{output_file}">DOSYAYI İNDİR</a>'

    return HTML


# -------------------------
# DOWNLOAD
# -------------------------
@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)


# -------------------------
# RUN
# -------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
