import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from docx import Document
import tempfile, os

PAGE_WIDTH, PAGE_HEIGHT = A4

# =========================
# DOCX
# =========================
def extract_docx(path):
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])

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
# PDF CREATE
# =========================
def create_content_pdf(text, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)

    TOP = 300
    BOTTOM = 70
    LEFT = 80

    FONT = "Times-Roman"
    SIZE = 12
    LINE = 16

    c.setFont(FONT, SIZE)

    y = PAGE_HEIGHT - TOP

    for line in text.split("\n"):
        line = line.strip()

        if not line:
            y -= LINE
            continue

        c.drawString(LEFT, y, line)
        y -= LINE

        if y < BOTTOM:
            c.showPage()
            c.setFont(FONT, SIZE)
            y = PAGE_HEIGHT - TOP

    c.save()

# =========================
# MERGE
# =========================
def merge_pdf(letterhead_pdf, content_pdf, output_pdf):
    writer = PdfWriter()

    letter = PdfReader(letterhead_pdf)
    content = PdfReader(content_pdf)

    for i in range(len(content.pages)):
        base = letter.pages[0]
        cont = content.pages[i]

        base.merge_page(cont)
        writer.add_page(base)

    with open(output_pdf, "wb") as f:
        writer.write(f)

def build_pdf(file_path, letterhead_pdf, output_pdf):
    content = get_content(file_path)

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name

    create_content_pdf(content, temp_file)
    merge_pdf(letterhead_pdf, temp_file, output_pdf)

    os.remove(temp_file)

# =========================
# UI FUNCTIONS
# =========================
def select_file():
    path = filedialog.askopenfilename(filetypes=[("Documents", "*.pdf;*.docx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, path)

def select_letterhead():
    path = filedialog.askopenfilename(filetypes=[("PDF Letterhead", "*.pdf")])
    letter_entry.delete(0, tk.END)
    letter_entry.insert(0, path)

def generate():
    file = file_entry.get()
    letter = letter_entry.get()

    if not file or not letter:
        messagebox.showerror("Error", "Select file and letterhead")
        return

    output = filedialog.asksaveasfilename(defaultextension=".pdf")

    try:
        build_pdf(file, letter, output)
        messagebox.showinfo("Success", "HEADEX Letter Generated!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# =========================
# MAIN WINDOW
# =========================
root = tk.Tk()
root.title("HEADEX - Center UI System")
root.geometry("520x420")
root.resizable(False, False)

# =========================
# BACKGROUND
# =========================
colors = ["#000814", "#001d3d", "#003566", "#001219"]
i = 0

def animate_bg():
    global i
    root.configure(bg=colors[i % len(colors)])
    i += 1
    root.after(1200, animate_bg)

# =========================
# CENTER CALC
# =========================
W = 520

# =========================
# UI (CENTERED)
# =========================

title = tk.Label(root,
                 text="HEADEX LETTERHEAD SYSTEM",
                 fg="white",
                 bg="#000814",
                 font=("Arial", 16, "bold"))
title.place(x=W/2, y=20, anchor="center")

file_entry = tk.Entry(root, width=40)
file_entry.place(x=W/2, y=90, anchor="center")

tk.Button(root, text="Browse File",
          command=select_file).place(x=W/2, y=120, anchor="center")

letter_entry = tk.Entry(root, width=40)
letter_entry.place(x=W/2, y=170, anchor="center")

tk.Button(root, text="Browse Letterhead",
          command=select_letterhead).place(x=W/2, y=200, anchor="center")

tk.Button(root,
          text="GENERATE LETTER",
          bg="#dc2626",
          fg="white",
          width=20,
          command=generate).place(x=W/2, y=260, anchor="center")

tk.Label(root,
         text="Powered by CODIXCO",
         fg="gray",
         bg="#000814").place(x=W/2, y=330, anchor="center")

# start animation
animate_bg()

root.mainloop()