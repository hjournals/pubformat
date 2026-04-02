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

HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Makale Sistemi</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 900px; margin: 30px auto; line-height: 1.5; }
        input, textarea { width: 100%; padding: 8px; margin-top: 4px; margin-bottom: 16px; box-sizing: border-box; }
        textarea { min-height: 120px; }
        button { padding: 10px 18px; font-size: 16px; }
    </style>
</head>
<body>
    <h1>Makale Oluştur</h1>

    <form method="POST" enctype="multipart/form-data">
        <label>Makale Başlığı</label>
        <input type="text" name="title" required>

        <label>Abstract</label>
        <textarea name="abstract" required></textarea>

        <label>Keywords</label>
        <input type="text" name="keywords" required>

        <label>Body (.docx)</label>
        <input type="file" name="body_file" accept=".docx" required>

        <label>Acknowledgements</label>
        <textarea name="acknowledgements"></textarea>

        <label>Funding</label>
        <textarea name="funding"></textarea>

        <label>Conflict of Interest</label>
        <textarea name="conflict_of_interest"></textarea>

        <label>References (her kaynağı ayrı satıra yapıştırın)</label>
        <textarea name="references_text" required></textarea>

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

def clean_reference_lines(text: str):
    lines = [line.strip() for line in text.splitlines()]
    return [line for line in lines if line]

def append_docx_content(target_doc, source_doc):
    for para in source_doc.paragraphs:
        new_p = target_doc.add_paragraph()

        # Stil adı varsa taşımayı dene
        try:
            if para.style and para.style.name:
                new_p.style = para.style.name
        except Exception:
            pass

        for run in para.runs:
            new_run = new_p.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline

            if run.font.name:
                new_run.font.name = run.font.name
            if run.font.size:
                new_run.font.size = run.font.size

        new_p.alignment = para.alignment

        src_fmt = para.paragraph_format
        dst_fmt = new_p.paragraph_format
        dst_fmt.left_indent = src_fmt.left_indent
        dst_fmt.right_indent = src_fmt.right_indent
        dst_fmt.first_line_indent = src_fmt.first_line_indent
        dst_fmt.space_before = src_fmt.space_before
        dst_fmt.space_after = src_fmt.space_after
        dst_fmt.line_spacing = src_fmt.line_spacing
        dst_fmt.line_spacing_rule = src_fmt.line_spacing_rule

def add_section_heading(doc, text):
    p = doc.add_paragraph()
    p.style = "Heading 1"
    p.add_run(text)

def add_normal_paragraph(doc, text):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p

def add_reference_paragraph(doc, text):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt = p.paragraph_format
    fmt.left_indent = Cm(1)
    fmt.first_line_indent = Cm(-1)
    fmt.space_before = Pt(6)
    fmt.space_after = Pt(6)
    fmt.line_spacing = 1
    return p

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            title = request.form.get("title", "").strip()
            abstract = request.form.get("abstract", "").strip()
            keywords = request.form.get("keywords", "").strip()
            acknowledgements = request.form.get("acknowledgements", "").strip()
            funding = request.form.get("funding", "").strip()
            conflict_of_interest = request.form.get("conflict_of_interest", "").strip()
            references_text = request.form.get("references_text", "").strip()
            blind_review = request.form.get("blind_review")

            author_block = "Anonymous Author(s)" if blind_review == "yes" else ""

            body_file = request.files.get("body_file")
            if not body_file or not body_file.filename.lower().endswith(".docx"):
                return "Lütfen body için .docx dosyası yükleyin."

            body_path = os.path.join(UPLOAD_FOLDER, body_file.filename)
            body_file.save(body_path)

            # 1) Şablonu doldur
            tpl = DocxTemplate(TEMPLATE_FILE)
            context = {
                "title": title,
                "author_block": author_block,
                "abstract": abstract,
                "keywords": keywords,
            }

            output_filename = "final_article.docx"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)

            tpl.render(context)
            tpl.save(output_path)

            # 2) Oluşan belgeyi aç
            final_doc = Document(output_path)

            # 3) Body'yi olduğu gibi sona ekle
            final_doc.add_page_break()
            body_doc = Document(body_path)
            append_docx_content(final_doc, body_doc)

            # 4) Acknowledgements
            if acknowledgements:
                add_section_heading(final_doc, "Acknowledgements")
                add_normal_paragraph(final_doc, acknowledgements)

            # 5) Funding
            if funding:
                add_section_heading(final_doc, "Funding")
                add_normal_paragraph(final_doc, funding)

            # 6) Conflict of Interest
            if conflict_of_interest:
                add_section_heading(final_doc, "Conflict of Interest")
                add_normal_paragraph(final_doc, conflict_of_interest)

            # 7) References
            add_section_heading(final_doc, "REFERENCES")
            for line in clean_reference_lines(references_text):
                add_reference_paragraph(final_doc, line)

            # 8) Kaydet
            final_doc.save(output_path)

            return render_template_string(RESULT_HTML, filename=output_filename)

        except Exception as e:
            return f"Hata oluştu: {str(e)}"

    return render_template_string(HTML_FORM)

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
