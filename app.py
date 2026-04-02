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
        file = request.files["file"]
        if file and file.filename.endswith(".docx"):
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            doc = Document(filepath)
paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

title = paragraphs[0] if len(paragraphs) > 0 else ""
author = paragraphs[1] if len(paragraphs) > 1 else ""

abstract = ""
body = []
keywords = ""

in_abstract = False

for p in paragraphs:
    if "abstract" in p.lower():
        in_abstract = True
        continue

    if "keywords" in p.lower():
        in_abstract = False
        keywords = p
        continue

    if in_abstract:
        abstract += p + " "
    else:
        body.append(p)

body_html = "<br><br>".join(body)

return f"""
<h1>{title}</h1>
<h3>{author}</h3>

<h2>Abstract</h2>
<p>{abstract}</p>

<h3>{keywords}</h3>

<hr>

<div>{body_html}</div>
"""

            return f"""
            <h2>Dosya okundu: {file.filename}</h2>
            <hr>
            <div>{content}</div>
            """

        return "Lütfen .docx dosyası yükleyin."

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
