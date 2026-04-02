from flask import Flask, request, render_template_string
import os
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <title>PubFormat</title>
</head>
<body>
    <h2>Makale Yükle</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Yükle</button>
    </form>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files.get("file")

        if not file:
            return "Dosya seçilmedi."

        if not file.filename.lower().endswith(".docx"):
            return "Lütfen .docx dosyası yükleyin."

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        doc = Document(filepath)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

        title = paragraphs[0] if len(paragraphs) > 0 else ""
        author = paragraphs[1] if len(paragraphs) > 1 else ""

        abstract = ""
        keywords = ""
        body = []
        in_abstract = False

        for p in paragraphs[2:]:
            lower_p = p.lower()

            if "abstract" in lower_p:
                in_abstract = True
                continue

            if "keywords" in lower_p:
                in_abstract = False
                keywords = p
                continue

            if in_abstract:
                abstract += p + " "
            else:
                body.append(p)

        body_html = "<br><br>".join(body)

        return f"""
<html>
<head>
    <title>Makale Önizleme</title>
    <style>
        body {{
            font-family: Georgia, serif;
            max-width: 800px;
            margin: auto;
            line-height: 1.6;
        }}
        h1 {{
            text-align: center;
            font-size: 28px;
        }}
        h3 {{
            text-align: center;
            font-weight: normal;
        }}
        h2 {{
            border-bottom: 1px solid #ccc;
            padding-bottom: 5px;
        }}
        .abstract {{
            background: #f5f5f5;
            padding: 10px;
            border-left: 4px solid #333;
        }}
        .keywords {{
            font-style: italic;
            margin-top: 10px;
        }}
    </style>
</head>

<body>

<h1>{title}</h1>
<h3>{author}</h3>

<hr>

<h2>Abstract</h2>
<div class="abstract">{abstract}</div>

<div class="keywords">{keywords}</div>

<hr>

<div>{body_html}</div>

<br><br>
<a href="/">← Yeni dosya yükle</a>

</body>
</html>
"""
        <html>
        <head>
            <title>Makale Önizleme</title>
        </head>
        <body>
            <h1>{title}</h1>
            <h3>{author}</h3>
            <hr>
            <h2>Abstract</h2>
            <p>{abstract}</p>
            <h3>{keywords}</h3>
            <hr>
            <div>{body_html}</div>
            <br><br>
            <a href="/">Yeni dosya yükle</a>
        </body>
        </html>
        """

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
