import os
import re
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image


def set_cell_margins(cell, top=150, start=150, bottom=150, end=150):
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


def insert_comparison_table(doc, before_images, after_images):
    all_keys = sorted(set(before_images.keys()) | set(after_images.keys()))

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    # Header
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Sebelum"
    hdr_cells[1].text = "Sesudah"
    for cell in hdr_cells:
        for para in cell.paragraphs:
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = para.runs[0]
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0, 0, 0)

    for key in all_keys:
        row_cells = table.add_row().cells

        # Sebelum
        if key in before_images:
            cell = row_cells[0]
            set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = p.add_run()
            try:
                with Image.open(before_images[key]) as im:
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
                    run.add_picture(before_images[key], width=width, height=height)
            except Exception as e:
                run.add_text(f"[Image not supported: {e}]")
        else:
            row_cells[0].text = "-"

        # Sesudah
        if key in after_images:
            cell = row_cells[1]
            set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = p.add_run()
            try:
                with Image.open(after_images[key]) as im:
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
                    run.add_picture(after_images[key], width=width, height=height)
            except Exception as e:
                run.add_text(f"[Image not supported: {e}]")
        else:
            row_cells[1].text = "-"

    doc.add_paragraph("")


def create_docx_comparison(folder_path, title1, title2="", output_filename="dokumentasi.docx"):
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

        # border under header
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
    before_images = {}
    after_images = {}

    for f in os.listdir(folder_path):
        ext = os.path.splitext(f)[1].lower()
        if ext not in valid_exts:
            continue
        match = re.match(r'([ba])(\d+)', f, re.IGNORECASE)
        if not match:
            continue
        prefix, num = match.groups()
        num = int(num)
        fp = os.path.join(folder_path, f)
        if prefix.lower() == "b":
            before_images[num] = fp
        elif prefix.lower() == "a":
            after_images[num] = fp

    if not (before_images or after_images):
        messagebox.showerror("Error", "Tidak ada gambar bX atau aX ditemukan.")
        return

    insert_comparison_table(doc, before_images, after_images)
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
    output_filename = os.path.basename(os.path.normpath(folder)) + "_comparison.docx"
    create_docx_comparison(folder, title1, title2, output_filename)


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
