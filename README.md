# Document Automation Suite

A Python desktop toolkit designed to automate photo report and table generation for procurement and civil engineering projects into Microsoft Word (`.docx`) documents.

## Key Features

- **Aspect-Ratio Preservation:** Automatically scales images up to 2.5 x 2.5 inches without distortion or overflow using Pillow.
- **OpenXML Table Formatting:** Applies custom cell margins, native paragraph borders, and `w:cantSplit` attributes so rows do not break across page boundaries.
- **Natural Sorting:** Uses regular expressions to sort alphanumeric filenames accurately (e.g., `b9` comes before `b10`).
- **Tkinter Desktop UI:** User-friendly file pickers and input forms designed for non-technical office staff.

## Included Tools

1. `DokGen1.2GUI.py`: Formats photos into a 2-column Word table with automatic filename captions.
2. `DokGen1.2NoCapGUI.py`: Formats photos into a clean 2-column Word table without text captions.
3. `DokGenSipil1GUI.py`: Sorts photos into three civil works stages: Before (`b*`), In-Progress (`p*`), and After (`a*`).
4. `DokGenSipilBnA1GUI.py`: Generates side-by-side Before and After (`b*` vs `a*`) comparison tables matching corresponding indices.

## Executables

Pre-compiled standalone `.exe` binaries for Windows are available under the **Releases** tab on GitHub, requiring no Python installation to run.

## Running from Source

Install dependencies:
```bash
pip install -r requirements.txt
```

Launch any tool:
```bash
python DokGen1.2GUI.py
```

## Compiling Standalone Executables

To package scripts into single executables using PyInstaller:
```bash
pyinstaller --onefile --windowed DokGen1.2GUI.py -n "Generator Dokumentasi Barang"
pyinstaller --onefile --windowed DokGen1.2NoCapGUI.py -n "Generator Dokumentasi Barang (No Caption)"
pyinstaller --onefile --windowed DokGenSipil1GUI.py -n "Generator Dokumentasi Sipil (Sebelum-Pengerjaan-Sesudah)"
pyinstaller --onefile --windowed DokGenSipilBnA1GUI.py -n "Generator Dokumentasi Sipil (Sebelum-Sesudah)"
```
