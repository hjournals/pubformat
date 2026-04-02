from flask import Flask, request, render_template_string, send_from_directory
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
TEMPLATE_FILE = "holistic template.docx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# -------------------------
# HTML
# -------------------------
HTML_FORM = """
<h2>Makale Sistemi</h2>

<form method="POST" enctype="multipart/form-data">

Başlık:<br>
<input name="title"><br><br>

Abstract:<br>
<textarea name="abstract"></textarea><br><br>

Keywords:<br>
<input name="keywords"><br><br>

Body (.docx):<br>
<input type="file" name="body_file"><br><br>

Acknowledgements:<br>
<textarea name="ack"></textarea><br><br>

Funding:<br>
<textarea name="funding"></textarea><br><br>

Conflict of Interest:<br>
<textarea name="conflict"></textarea><br><br>

References:<br>
<textarea name="references"></textarea><br><br>

<input type="checkbox" name="blind"> Kör hakem<br><br>

<button>Oluştur</button>

</form>
"""

# -------------------------
# BODY EKLEME (OLDUĞU GİBİ)
# -------------------------
def append_body(target_doc, source_doc):
    for para in source_doc.paragraphs:
        new_p = target_doc.add_paragraph()

        for run in para.runs:
            new_run = new_p.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline

        new_p.alignment = para.alignment

# -------------------------
# REFERENCES FORMAT
# -------------------------
def add_references(doc, text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    for line in lines:
        p = doc.add_paragraph(line)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        fmt = p.paragraph_format
        fmt.left_indent = Cm(1)
        fmt.first_line_indent = Cm(-1)
        fmt.space_before = Pt(6)
        fmt.space_after = Pt(6)

# -------------------------
# ROUTE
# -------------------------
@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":
        try:
            title = request.form.get("title")
            abstract = request.form.get("abstract")
            keywords = request.form.get("keywords")
            ack = request.form.get("ack")
            funding = request.form.get("funding")
            conflict = request.form.get("conflict")
            references = request.form.get("references")

            author_block = "Anonymous Author(s)" if request.form.get("blind") else ""

            # BODY FILE
            file = request.files["body_file"]
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            # TEMPLATE RENDER
            doc = DocxTemplate(TEMPLATE_FILE)

            context = {
                "title": title,
                "author_block": author_block,
                "abstract": abstract,
                "keywords": keywords,
            }

            output_file = "final.docx"
            output_path = os.path.join(OUTPUT_FOLDER, output_file)

            doc.render(context)
            doc.save(output_path)

            # FINAL DOC AÇ
            final_doc = Document(output_path)

            # BODY EKLE
            final_doc.add_page_break()
            body_doc = Document(filepath)
            append_body(final_doc, body_doc)

            # ACKNOWLEDGEMENTS
            final_doc.add_paragraph()
            p = final_doc.add_paragraph("Acknowledgements")
            p.runs[0].bold = True
            final_doc.add_paragraph(ack)

            # FUNDING
            p = final_doc.add_paragraph("Funding")
            p.runs[0].bold = True
            final_doc.add_paragraph(funding)

            # CONFLICT
            p = final_doc.add_paragraph("Conflict of Interest")
            p.runs[0].bold = True
            final_doc.add_paragraph(conflict)

            # REFERENCES
            p = final_doc.add_paragraph("REFERENCES")
            p.runs[0].bold = True
            add_references(final_doc, references)

            # SAVE
            final_doc.save(output_path)

            return f'<a href="/download/{output_file}">İndir</a>'

        except Exception as e:
            return f"HATA: {str(e)}"

    return HTML_FORM

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
