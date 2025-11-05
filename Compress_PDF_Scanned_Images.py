#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import fitz  # PyMuPDF
from PIL import Image
import io, os, PySimpleGUI as sg
import threading
import sys

# ---------- PDF COMPRESSION ENGINE (unchanged) ----------
def compress_pdf(input_path, output_path, dpi, jpeg_quality):
    """Handles compression for a single PDF (Used repeatedly in batch mode)."""
    doc = fitz.open(input_path)
    images = []
    total_pages = len(doc)

    for i in range(total_pages):
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = doc.load_page(i).get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("ppm")))
        img_io = io.BytesIO()
        img.save(img_io, format="JPEG", quality=jpeg_quality, optimize=True)
        compressed_img = Image.open(io.BytesIO(img_io.getvalue())).convert("RGB")
        images.append(compressed_img)

    if images:
        images[0].save(
            output_path,
            save_all=True,
            append_images=images[1:],
            resolution=dpi
        )

    original_size = os.path.getsize(input_path) / 1024 / 1024
    new_size = os.path.getsize(output_path) / 1024 / 1024
    return (os.path.basename(input_path), original_size, new_size)


def process_pdfs(pdfs, input_folder, output_folder, dpi, jpeg_quality, window):
    total = len(pdfs)
    for idx, pdf in enumerate(pdfs, 1):
        try:
            in_file = os.path.join(input_folder, pdf)
            out_file = os.path.join(output_folder, os.path.splitext(pdf)[0] + "_compressed.pdf")
            name, orig, comp = compress_pdf(in_file, out_file, dpi, jpeg_quality)
            window.write_event_value("-UPDATE-", (idx, total, f"✅ {name}: {orig:.2f} → {comp:.2f} MB"))
        except Exception as e:
            window.write_event_value("-UPDATE-", (idx, total, f"❌ {pdf}: {e}"))
    window.write_event_value("-DONE-", None)

# ---------- BRANDING / THEME ----------
PRIMARY = "#007C92"   # teal
ACCENT  = "#00B8A9"   # light teal
BG      = "#121212"   # dark charcoal
PANEL   = "#1E1E1E"
INK     = "#EAEAEA"   # light text
INK_SUB = "#B9C4C9"
FIELD   = "#242424"
BORDER  = "#333333"

if getattr(sys, 'frozen', False):
    # Running as a compiled exe
    BASE_PATH = sys._MEIPASS
else:
    # Running as a normal script
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

LOGO_PATH = os.path.join(BASE_PATH, "KIEN_LOGO_SMALL.png")
ICON_PATH = os.path.join(BASE_PATH, "KIEN_ICON.ico")


sg.theme_add_new("KienDark", {
    "BACKGROUND": BG,
    "TEXT": INK,
    "INPUT": FIELD,
    "TEXT_INPUT": INK,
    "SCROLL": FIELD,
    "BUTTON": (INK, PRIMARY),
    "PROGRESS": (ACCENT, FIELD),
    "BORDER": 1,
    "SLIDER_DEPTH": 0,
    "PROGRESS_DEPTH": 0
})
sg.theme("KienDark")
sg.set_options(font=("Segoe UI", 10), input_elements_background_color=FIELD)

# ---------- Helper: load & scale logo to fixed height WITHOUT cropping ----------
def logo_bytes(path, target_h=40):
    im = Image.open(path).convert("RGBA")
    scale = target_h / float(im.height)
    new_w = int(round(im.width * scale))
    im = im.resize((new_w, target_h), Image.LANCZOS)
    bio = io.BytesIO()
    im.save(bio, format="PNG")
    return bio.getvalue()

# ---------- HEADER ----------
header = [[
    sg.Image(data=logo_bytes(LOGO_PATH, 40), pad=((10,10),(8,8)), background_color=BG),
    sg.Text("PDF Compressor", font=("Segoe UI", 18, "bold"), text_color=INK, background_color=BG, pad=(0,0)),
    sg.Text("  by Kien Consultants", font=("Segoe UI", 11), text_color=ACCENT, background_color=BG, pad=(0,0)),
]]

# ---------- CONTROLS ----------
modes = [[
    sg.Radio("Single File", "MODE", default=True, key="-SINGLE-", text_color=INK, background_color=BG),
    sg.Radio("Batch Folder", "MODE", key="-BATCH-", text_color=INK, background_color=BG)
]]

controls = [
    [sg.Text("Select PDF File or Folder:", background_color=BG, text_color=INK_SUB),
     sg.Input(key="-INPUT-", size=(58,1), background_color=FIELD, text_color=INK, border_width=1),
     sg.Button("Browse", key="-BROWSE-", button_color=(INK, PRIMARY))],

    [sg.Text("Output Folder (optional):", background_color=BG, text_color=INK_SUB),
     sg.Input(key="-OUTPUT-", size=(58,1), background_color=FIELD, text_color=INK, border_width=1),
     sg.FolderBrowse(button_text="Browse", button_color=(INK, PRIMARY))],

    [sg.Text("DPI:", size=(6,1), text_color=INK_SUB, background_color=BG),
     sg.Slider(range=(72,300), default_value=120, orientation="h", key="-DPI-", s=(42,15),
               trough_color=FIELD, relief=sg.RELIEF_FLAT),
     sg.Text("JPEG Quality:", size=(12,1), text_color=INK_SUB, background_color=BG, pad=((14,0),0)),
     sg.Slider(range=(10,100), default_value=50, orientation="h", key="-QUAL-", s=(42,15),
               trough_color=FIELD, relief=sg.RELIEF_FLAT)],
]

# ---------- STATUS + PROGRESS (no extra blank row) ----------
status_row = [[
    sg.Text("Status:", text_color=INK_SUB, background_color=BG),
    sg.Text("", key="-STATUS-", size=(55,1), text_color=INK, background_color=PANEL),
    sg.ProgressBar(100, orientation="h", size=(25,18), key="-PROG-", bar_color=(ACCENT, FIELD),
                   border_width=0, visible=False)
]]

# ---------- LOG PANEL ----------
log_panel = [[
    sg.Multiline(size=(90,18), key="-LOG-", autoscroll=True, background_color=FIELD,
                 text_color=INK, border_width=1)
]]

# ---------- ACTIONS ----------
actions = [[
    sg.Button("Start Compression", key="Start Compression", size=(20,1), button_color=(INK, PRIMARY)),
    sg.Button("Exit", size=(10,1), button_color=(INK, "#3A3A3A"))
]]

# ---------- LAYOUT ----------
layout = [
    [sg.Column(header, background_color=BG, pad=(6,6))],
    [sg.HorizontalSeparator(color=BORDER)],
    [sg.Column(modes,     background_color=BG, pad=((6,6),(6,2)))],
    [sg.Column(controls,  background_color=BG, pad=((6,6),(0,8)))],
    [sg.Column(status_row,background_color=BG, pad=((6,6),(0,6)))],
    [sg.Column(log_panel, background_color=BG, pad=((6,6),(0,6)))],
    [sg.Column(actions,   background_color=BG, pad=((6,6),(4,10)))]
]

window = sg.Window(
    "PDF Compressor by Kien Consultants",
    layout,
    icon=ICON_PATH,
    background_color=BG,
    finalize=True
)

# ---------- EVENT LOOP (functional logic unchanged) ----------
while True:
    event, values = window.read(timeout=200)

    if event in (sg.WINDOW_CLOSED, "Exit"):
        break

    # Dynamic browse: single file vs folder
    if event == "-BROWSE-":
        if values.get("-SINGLE-", True):
            file_path = sg.popup_get_file(
                "Select a PDF file",
                file_types=(("PDF Files", "*.pdf"),),
                no_window=True
            )
            if file_path:
                window["-INPUT-"].update(file_path)
        else:
            folder_path = sg.popup_get_folder("Select a folder containing PDFs", no_window=True)
            if folder_path:
                window["-INPUT-"].update(folder_path)
        continue

    if event == "Start Compression":
        dpi = int(values["-DPI-"])
        jpeg_quality = int(values["-QUAL-"])
        input_path = values["-INPUT-"]
        output_folder = values["-OUTPUT-"]

        if not input_path or not os.path.exists(input_path):
            sg.popup_error("Please select a valid PDF file or folder.")
            continue

        if os.path.isfile(input_path) and input_path.lower().endswith(".pdf"):
            output_path = output_folder or os.path.dirname(input_path)
            output_file = os.path.join(
                output_path,
                os.path.splitext(os.path.basename(input_path))[0] + "_compressed.pdf"
            )
            window["-STATUS-"].update("Compressing single file...")
            window["-PROG-"].update(0)
            window["-PROG-"].update_bar(0)
            window["-PROG-"].update(visible=True)
            try:
                name, orig, comp = compress_pdf(input_path, output_file, dpi, jpeg_quality)
                window["-LOG-"].print(f"✅ {name}: {orig:.2f} → {comp:.2f} MB")
                window["-STATUS-"].update("✅ Completed!")
            except Exception as e:
                window["-LOG-"].print(f"❌ Error: {e}")
                window["-STATUS-"].update("❌ Error during compression")
            finally:
                window["-PROG-"].update(visible=False)

        elif os.path.isdir(input_path):
            pdfs = [f for f in os.listdir(input_path) if f.lower().endswith(".pdf")]
            if not pdfs:
                sg.popup_error("No PDF files found in the selected folder.")
                continue

            window["-STATUS-"].update(f"Compressing {len(pdfs)} files...")
            window["-LOG-"].print(f"📁 Found {len(pdfs)} PDFs in {input_path}\n")
            window["-PROG-"].update(0)
            window["-PROG-"].update(visible=True)

            threading.Thread(
                target=process_pdfs,
                args=(pdfs, input_path, output_folder or input_path, dpi, jpeg_quality, window),
                daemon=True
            ).start()
        else:
            sg.popup_error("Invalid selection. Please choose a PDF file or a folder containing PDFs.")
            continue

    elif event == "-UPDATE-":
        idx, total, msg = values[event]
        window["-PROG-"].update((idx / total) * 100)
        window["-STATUS-"].update(f"{idx}/{total} processed")
        window["-LOG-"].print(msg)

    elif event == "-DONE-":
        window["-PROG-"].update(visible=False)
        window["-STATUS-"].update("🏁 All files processed!")
        window["-LOG-"].print("\n✅ Batch compression complete!\n")

window.close()
