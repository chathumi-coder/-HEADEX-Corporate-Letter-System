from flask import Flask, render_template, request, send_file
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
import os
from copy import deepcopy

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ================= DOCX READ =================
def extract_docx(path):
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])

# ================= GET CONTENT =================
def get_content(file_path):
    if file_path.lower().endswith(".pdf"):
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n"
        return text

    elif file_path.lower().endswith(".docx"):
        return extract_docx(file_path)

    return ""

# ================= CREATE CONTENT PDF =================
def create_pdf(text, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    left_margin = 80
    right_margin = 80
    top_margin = 250
    bottom_margin = 100

    max_width = width - left_margin - right_margin

    c.setFont("Helvetica", 11)
    y = height - top_margin

    for line in text.split("\n"):
        words = line.split(" ")
        current_line = ""

        for word in words:
            test_line = current_line + word + " "
            if c.stringWidth(test_line, "Helvetica", 11) < max_width:
                current_line = test_line
            else:
                c.drawString(left_margin, y, current_line)
                y -= 18
                current_line = word + " "

                if y < bottom_margin:
                    c.showPage()
                    c.setFont("Helvetica", 11)
                    y = height - top_margin

        if current_line:
            c.drawString(left_margin, y, current_line)
            y -= 18

            if y < bottom_margin:
                c.showPage()
                c.setFont("Helvetica", 11)
                y = height - top_margin

    c.save()

# ================= MERGE LETTERHEAD =================
def merge_pdf(letter_path, content_path, output_path):
    writer = PdfWriter()

    letter_pdf = PdfReader(letter_path)
    content_pdf = PdfReader(content_path)

    pages = max(len(letter_pdf.pages), len(content_pdf.pages))

    for i in range(pages):
        base = letter_pdf.pages[i] if i < len(letter_pdf.pages) else letter_pdf.pages[0]
        page = deepcopy(base)

        if i < len(content_pdf.pages):
            page.merge_page(content_pdf.pages[i])

        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)

# ================= BUILD =================
def build(file_path, letter_path):
    text = get_content(file_path)

    temp_pdf = os.path.join(OUTPUT_FOLDER, "content.pdf")
    final_pdf = os.path.join(OUTPUT_FOLDER, "final.pdf")

    create_pdf(text, temp_pdf)
    merge_pdf(letter_path, temp_pdf, final_pdf)

    return final_pdf

# ================= ROUTES =================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    if "file" not in request.files or "letter" not in request.files:
        return "Missing files"

    file = request.files["file"]
    letter = request.files["letter"]

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    letter_path = os.path.join(UPLOAD_FOLDER, letter.filename)

    file.save(file_path)
    letter.save(letter_path)

    final_pdf = build(file_path, letter_path)

    return send_file(final_pdf, as_attachment=True)

# ================= RUN =================
if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
