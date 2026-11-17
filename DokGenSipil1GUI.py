import os
import re
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image


def set_cell_margins(cell, top=100, start=100, bottom=100, end=100):
    """Set cell margins in twips (1/1440 of an inch)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{margin}")
        node.set(qn("w:type"), "dxa")
        node.set(qn("w:w"), str(value))
        tcMar.append(node)
    tcPr.append(tcMar)


def natural_sort_key(filename):
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else 0


def insert_image_table(doc, images, section_title):
    if not images:
        return

    # Section title centered
    para = doc.add_paragraph()
    para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    run = para.add_run(section_title)
    run.font.size = Pt(12)
    run.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)

    # Table
    rows = (len(images) + 1) // 2
    table = doc.add_table(rows=rows, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    for row in table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        cantSplit = OxmlElement("w:cantSplit")
        trPr.append(cantSplit)

    for i, img_path in enumerate(images):
        row = i // 2
        col = i % 2
        cell = table.cell(row, col)
        set_cell_margins(cell, top=150, start=150, bottom=150, end=150)

        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run()
        try:
            with Image.open(img_path) as im:
                width_px, height_px = im.size
                aspect = width_px / height_px
                max_width = Inches(2.5)
                max_height = Inches(2.5)
                if aspect >= 1:
                    width = min(max_width, Inches(width_px / 96))
                    height = width / aspect
                else:
                    height = min(max_height, Inches(height_px / 96))
                    width = height * aspect
                run.add_picture(img_path, width=width, height=height)
        except Exception as e:
            run.add_text(f"[Image not supported: {e}]")
    doc.add_paragraph("")  # spacing after table


def create_docx_from_images(folder_path, title1, title2="", output_filename="dokumentasi.docx"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

        header = section.header
        header_paragraph = header.paragraphs[0]
        header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        full_title = title1
        if title2.strip():
            full_title += "\n" + title2.strip()

        run = header_paragraph.add_run(full_title)
        run.bold = True
        run.font.size = Pt(14)

        p = header_paragraph._element
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), "000000")
        pBdr.append(bottom)
        pPr.append(pBdr)
        header_paragraph.paragraph_format.space_after = Pt(18)

    valid_exts = [".jpg", ".jpeg", ".png", ".bmp", ".gif"]
    before, process, after = [], [], []

    files = [f for f in os.listdir(folder_path) if os.path.splitext(f)[1].lower() in valid_exts]
    files.sort(key=natural_sort_key)

    for f in files:
        fp = os.path.join(folder_path, f)
        if f.lower().startswith("b"):
            before.append(fp)
        elif f.lower().startswith("p"):
            process.append(fp)
        elif f.lower().startswith("a"):
            after.append(fp)

    if not (before or process or after):
        messagebox.showerror("Error", "Tidak ada gambar ditemukan dalam folder.")
        return

    insert_image_table(doc, before, "Sebelum")
    insert_image_table(doc, process, "Pengerjaan")
    insert_image_table(doc, after, "Sesudah")

    doc.save(output_filename)
    messagebox.showinfo("Sukses", f"Dokumentasi berhasil dibuat:\n{output_filename}")


# --- GUI ---
def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_entry.delete(0, "end")
        folder_entry.insert(0, folder)


def generate_docx():
    folder = folder_entry.get()
    title1 = title1_entry.get()
    title2 = title2_entry.get()
    if not os.path.isdir(folder):
        messagebox.showerror("Error", "Alamat folder tidak valid.")
        return
    output_filename = os.path.basename(os.path.normpath(folder)) + ".docx"
    create_docx_from_images(folder, title1, title2, output_filename)


root = Tk()
root.title("Generator Dokumentasi Sipil")
root.geometry("500x250")

Label(root, text="Pilih folder gambar:").pack(pady=(10,0))
folder_entry = Entry(root, width=60)
folder_entry.pack(pady=5)
Button(root, text="Browse...", command=browse_folder).pack(pady=(0,10))

Label(root, text="Judul 1:").pack()
title1_entry = Entry(root, width=60)
title1_entry.pack(pady=5)

Label(root, text="Judul 2 (opsional):").pack()
title2_entry = Entry(root, width=60)
title2_entry.pack(pady=5)

Button(root, text="Buat Dokumentasi", command=generate_docx, bg="green", fg="white").pack(pady=15)

root.mainloop()
