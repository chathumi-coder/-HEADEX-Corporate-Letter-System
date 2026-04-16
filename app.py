from flask import Flask, request, send_file
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
import tempfile, os, uuid
from copy import deepcopy

app = Flask(__name__)

# =========================
# DOCX READER
# =========================
def extract_docx(path):
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])

# =========================
# GET TEXT
# =========================
def get_content(file_path):
    if file_path.endswith(".pdf"):
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n"
        return text

    elif file_path.endswith(".docx"):
        return extract_docx(file_path)

    return ""

# =========================
# CREATE TEXT PDF
# =========================
def create_pdf(text, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)

    TOP = 750
    BOTTOM = 80
    LEFT = 80

    c.setFont("Helvetica", 11)
    y = TOP

    for line in text.split("\n"):
        line = line.strip()

        if line:
            c.drawString(LEFT, y, line[:110])
            y -= 15

            if y < BOTTOM:
                c.showPage()
                c.setFont("Helvetica", 11)
                y = TOP

    c.save()

# =========================
# MERGE LETTERHEAD + TEXT
# =========================
def merge_pdf(letter_pdf, content_pdf, output_pdf):
    writer = PdfWriter()

    letter = PdfReader(letter_pdf)
    content = PdfReader(content_pdf)

    pages = max(len(letter.pages), len(content.pages))

    for i in range(pages):
        base = letter.pages[i] if i < len(letter.pages) else letter.pages[0]
        new_page = deepcopy(base)

        if i < len(content.pages):
            new_page.merge_page(content.pages[i])

        writer.add_page(new_page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

# =========================
# BUILD FINAL PDF
# =========================
def build(file_path, letter_path):
    text = get_content(file_path)

    temp_text = os.path.join(tempfile.gettempdir(), f"text_{uuid.uuid4()}.pdf")
    final_pdf = os.path.join(tempfile.gettempdir(), f"final_{uuid.uuid4()}.pdf")

    create_pdf(text, temp_text)
    merge_pdf(letter_path, temp_text, final_pdf)

    return final_pdf

# =========================
# HOME ROUTE (ONLINE SAFE)
# =========================
@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":
        file = request.files.get("file")
        letter = request.files.get("letter")

        if not file or not letter:
            return "❌ Upload both file + letterhead"

        file_path = os.path.join(tempfile.gettempdir(), file.filename)
        letter_path = os.path.join(tempfile.gettempdir(), letter.filename)

        file.save(file_path)
        letter.save(letter_path)

        output = build(file_path, letter_path)

        # 👉 DIRECT OPEN PDF (NO PREVIEW PAGE)
        return send_file(output, mimetype="application/pdf")

    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>HEADEX ONLINE</title>
        <style>
            body{
                background:#0f172a;
                color:white;
                text-align:center;
                font-family:Arial;
                padding:50px;
            }
            .box{
                background:#1e293b;
                padding:30px;
                border-radius:10px;
                width:400px;
                margin:auto;
            }
            input,button{
                margin:10px;
            }
            button{
                padding:10px 20px;
                background:#ef4444;
                color:white;
                border:none;
                cursor:pointer;
            }
        </style>
    </head>
    <body>

        <div class="box">
            <h2>HEADEX ONLINE</h2>

            <form method="POST" enctype="multipart/form-data">
                <p>Upload Document</p>
                <input type="file" name="file" required>

                <p>Upload Letterhead PDF</p>
                <input type="file" name="letter" required>

                <br>
                <button type="submit">GENERATE & OPEN PDF</button>
            </form>
        </div>

    </body>
    </html>
    '''

# =========================
# RUN (RENDER / RAILWAY SUPPORT)
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
