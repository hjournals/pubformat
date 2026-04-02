from flask import Flask, request, render_template_string, send_from_directory
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import re

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
        h1 { margin-bottom: 24px; }
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

def find_paragraph_index(document, marker_text):
    for i, p in enumerate(document.paragraphs):
        if marker_text in p.text:
            return i
    return None

def insert_paragraph_after(paragraph, text=""):
    new_p = paragraph._p.addnext(paragraph._element.__class__())
    from docx.text.paragraph import Paragraph
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para

def copy_body_docx_into_document(target_doc, body_doc):
    for para in body_doc.paragraphs:
        new_p = target_doc.add_paragraph()
        new_p.style = para.style

        for run in para.runs:
            new_run = new_p.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            if run.font.name:
                new_run.font.name = run.font.name
            if run.font.size:
                new_run.font.size = run.font.size

        # hizalama
        new_p.alignment = para.alignment

        # paragraf biçimi
        pf_src = para.paragraph_format
        pf_dst = new_p.paragraph_format

        pf_dst.left_indent = pf_src.left_indent
        pf_dst.right_indent = pf_src.right_indent
        pf_dst.first_line_indent = pf_src.first_line_indent
        pf_dst.keep_together = pf_src.keep_together
        pf_dst.keep_with_next = pf_src.keep_with_next
        pf_dst.page_break_before = pf_src.page_break_before
        pf_dst.widow_control = pf_src.widow_control
        pf_dst.space_before = pf_src.space_before
        pf_dst.space_after = pf_src.space_after
        pf_dst.line_spacing = pf_src.line_spacing
        pf_dst.line_spacing_rule = pf_src.line_spacing_rule

def add_references_with_formatting(document, marker_text, references_text):
    idx = find_paragraph_index(document, marker_text)
    if idx is None:
        return

    marker_paragraph = document.paragraphs[idx]
    lines = clean_reference_lines(references_text)

    for line in lines:
        new_p = insert_paragraph_after(marker_paragraph, line)

        # Word biçimi
        new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        fmt = new_p.paragraph_format
        fmt.left_indent = Cm(0)
        fmt.right_indent = Cm(0)
        fmt.first_line_indent = Cm(-1)   # asılı girinti
        fmt.left_indent = Cm(1)          # 1 cm içerden başlasın
        fmt.space_before = Pt(6)
        fmt.space_after = Pt(6)
        fmt.line_spacing = 1

        marker_paragraph = new_p

    # marker satırını temizle
    document.paragraphs[idx].text = ""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
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

        # 1) Şablonu render et
        tpl = DocxTemplate(TEMPLATE_FILE)
        context = {
            "title": title,
            "author_block": author_block,
            "abstract": abstract,
            "keywords": keywords,
            "acknowledgements": acknowledgements,
            "funding": funding,
            "conflict_of_interest": conflict_of_interest,
            "references_marker": "{{references_marker}}"
        }

        temp_rendered_path = os.path.join(OUTPUT_FOLDER, "temp_rendered.docx")
        tpl.render(context)
        tpl.save(temp_rendered_path)

        # 2) Render edilmiş dosyayı aç
        final_doc = Document(temp_rendered_path)

        # 3) Body dosyasını aynen ekle
        body_doc = Document(body_path)
        final_doc.add_page_break()
        copy_body_docx_into_document(final_doc, body_doc)

        # 4) References marker altına kaynakçayı biçimlendirerek ekle
        add_references_with_formatting(final_doc, "{{references_marker}}", references_text)

        # 5) Kaydet
        output_filename = "final_article.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        final_doc.save(output_path)

        return render_template_string(RESULT_HTML, filename=output_filename)

    return render_template_string(HTML_FORM)

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
