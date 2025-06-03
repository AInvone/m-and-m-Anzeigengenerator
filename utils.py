import os
import re
from datetime import datetime
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from openai import OpenAI

# ---------- LLM Functions ----------

def generate_real_estate_ad_deeps(args, prompt):
    client = OpenAI(
        api_key=args['deepseek_api_key'],
        base_url=args['deepseek_base_url'],
    )
    response = client.chat.completions.create(
        model=args['deepseek_model'],
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

def create_prompt(params):
    return f"""Erstelle eine professionelle Immobilienanzeige für Immoscout24 auf Deutsch mit diesen Angaben:
Aktion: {params['action']}
Immobilientyp: {params['immotype']}
Ort: {params['location']}
Preis: {params['price']}
Wohnfläche: {params['living_area']}
Grundstücksgröße: {params['lot_size']}
Zimmeranzahl: {params['rooms']}
Baujahr: {params['year_built']}
Haustyp: {params['house_type']}
Heizungsart: {params['heating']}
Internetgeschwindigkeit: {params['internet_speed']}
Energieeffizienzklasse: {params['energy_efficiency_class']}
Ausstattung/Merkmale: {params['features']}

Bitte verwende eine übersichtliche Struktur: 
- Eine ansprechende und aufmerksamkeitsstarke Überschrift
- Einen Abschnitt "Key Facts" mit den wichtigsten Eckdaten
- Eine Merkmals-Tabelle (in Markdown-Tabelle mit zwei Spalten: Merkmal | Wert)
- Einen langen, ausführlichen, umfangreichen und ansprechenden Haupttext in Fließtext mit mindestens 4 Abschnitten mit Hervorhebung aller Vorteile (USPs)
- Die Fließtext-Absätze sollen länger und ausgeschweifter sein, bitte rethorisch ausgeschmückt, oder auch mit einem oder zwei inhaltlichen "Hooks", bevor es zu "Ihre Vorteile auf einen Blick:" kommt
- Zielgruppe: Käufer auf dem deutschen Immobilienmarkt, die gerne längere Texte lesen
- Stil: inspirierend, klar, realistisch, ansprechend, geduldig viel erzählend.

Gib nur den Anzeigentext zurück und kommentiere nicht die obigen Anweisungen.
"""

def trim_trailing_notes(text: str) -> str:
    lines = text.strip().split("\n")
    while lines and (
        "diese anzeige" in lines[-1].lower()
        or lines[-1].strip() in ("---", "")
    ):
        lines.pop()
    return "\n".join(lines)

# ---------- Markdown Cleaning ----------

def _process_markdown_line(line):
    parts = []
    buffer = []
    in_bold = False
    for char in line:
        if char == "*" and len(buffer) >= 1 and buffer[-1] == "*":
            buffer.pop()
            if in_bold:
                parts.append(('bold', ''.join(buffer)))
                in_bold = False
            else:
                parts.append(('text', ''.join(buffer)))
                in_bold = True
            buffer = []
        else:
            buffer.append(char)
    if buffer:
        parts.append(('text', ''.join(buffer)))
    return parts

def clean_markdown(text):
    lines = []
    first_line = True

    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue

        # HEADING 1: erste fette Zeile als Haupttitel
        if first_line and line.startswith("**") and line.endswith("**"):
            clean_title = line.strip("*").strip()
            lines.append(("heading1", [("text", clean_title)]))
            first_line = False
            continue
        first_line = False

        # Markdown-Tabelle
        if "|" in line and re.match(r"^\|.*\|$", line):
#            cells = [re.sub(r"\*+", "", c.strip()) for c in line.strip().strip("|").split("|")]
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            lines.append(("table_row", cells))
            continue

        # Unterüberschriften
        if line.startswith("###"):
            lines.append(("heading2", line.replace("###", "").strip()))
            continue
        elif line.startswith("**") and line.endswith("**"):
            line_clean = re.sub(r"\*+", "", line).strip()
            lines.append(("heading3", line_clean))
            continue

        # Bullet (aber keine kaputten wie • --)
        bullet_match = re.match(r'^\s*[-•✔🔹]\s+(.+)', line)
        if bullet_match and bullet_match.group(1).strip() != "--":
            lines.append(("bullet", (1, _process_markdown_line(bullet_match.group(1)))))
            continue

        # Fließtext
        processed = _process_markdown_line(line)
        if any(t[0] == 'bold' for t in processed):
            lines.append(("bold_text", processed))
        else:
            lines.append(("paragraph", processed))

    return lines


def set_table_border(tbl, border_dir, val="single", size="4", color="888888"):
    """
    Setzt Tabellenrahmen in einem docx-Tabellelement.
    """
    tbl_pr = tbl.tblPr
    tbl_borders = tbl_pr.tblBorders or OxmlElement("w:tblBorders")

    border = OxmlElement(f"w:{border_dir}")
    border.set(qn("w:val"), val)
    border.set(qn("w:sz"), size)
    border.set(qn("w:space"), "0")
    border.set(qn("w:color"), color)

    tbl_borders.append(border)
    tbl_pr.append(tbl_borders)


def add_clean_table_to_docx(doc, rows, bold_header=True):
    """
    Fügt eine schön formatierte Tabelle zu einem Word-Dokument hinzu.
    Args:
        doc: docx.Document Objekt
        rows: Liste von Zeilen, jede Zeile ist Liste von Zellen (strings)
        bold_header: Ob erste Zeile fett formatiert wird
    """
    if not rows or not all(isinstance(r, list) for r in rows):
        return

    # Zeile mit nur Bindestrichen herausfiltern
    filtered_rows = [
        r for r in rows
        if not all(cell.strip().startswith("---") or set(cell.strip()) == {"-"} for cell in r)
    ]

    if not filtered_rows:
        return

    table = doc.add_table(rows=0, cols=len(filtered_rows[0]))
    table.style = "Table Grid"

    # Tabellenlinien: Dunkelgrau
    tbl = table._tbl
    for border_dir in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        set_table_border(tbl, border_dir, "single", "4", "888888")  # 888888 = dunkelgrau

    for row_idx, row_cells in enumerate(filtered_rows):
        row = table.add_row()
        for col_idx, cell_text in enumerate(row_cells):
            clean_text = cell_text.replace("**", "").strip()
            cell = row.cells[col_idx]
            para = cell.paragraphs[0]
            run = para.add_run(clean_text)

            if bold_header and row_idx == 0:
                run.bold = True
            elif "**" in cell_text:
                run.bold = True

            para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    return table


# ---------- Image Helper ----------

def add_logo_top_right(doc, logo_path):
    if logo_path and os.path.exists(logo_path):
        section = doc.sections[0]
        section.top_margin = Inches(1.2)
        header = section.header
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        p.add_run().add_picture(logo_path, width=Inches(2))

def rescale_img(input_path_or_folder, output_folder="data/rescaled_details", max_width=1500, quality=100):
    """
    Rescales a single image or all images in a folder to max width.
    Saves them as optimized JPEGs in `output_folder` and returns paths.

    Args:
        input_path_or_folder (str): Path to image file or directory.
        output_folder (str): Where to save rescaled images.
        max_width (int): Max pixel width.
        quality (int): JPEG quality (default 85).

    Returns:
        str | dict: Rescaled image path (if single image) or dict of {original_path: resized_path}
    """
    os.makedirs(output_folder, exist_ok=True)

    def _rescale_and_save(img_path):
        try:
            with Image.open(img_path) as img:
                img_format = img.format or "JPEG"
                width, height = img.size

                if width > max_width:
                    new_height = int(max_width * height / width)
                    img = img.resize((max_width, new_height), Image.LANCZOS)

                base_name = os.path.splitext(os.path.basename(img_path))[0] + ".jpg"
                out_path = os.path.join(output_folder, base_name)
                img.convert("RGB").save(out_path, format="JPEG", quality=quality, optimize=True)
                return out_path
        except Exception as e:
            print(f"⚠️ Fehler beim Skalieren von {img_path}: {e}")
            return img_path  # fallback

    if os.path.isfile(input_path_or_folder):
        return _rescale_and_save(input_path_or_folder)
    
    elif os.path.isdir(input_path_or_folder):
        results = {}
        for f in os.listdir(input_path_or_folder):
            full_path = os.path.join(input_path_or_folder, f)
            if f.lower().endswith((".png", ".jpg", ".jpeg")):
                resized = _rescale_and_save(full_path)
                results[full_path] = resized
        return results

    else:
        raise ValueError(f"Pfad nicht gefunden: {input_path_or_folder}")
    
def add_main_image(doc, path, caption=""):
    if path and os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run().add_picture(path, width=Inches(5.5))
        if caption:
            caption_paragraph = doc.add_paragraph(caption, style="Caption")
            caption_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def clean_filename_for_caption(fname):
    name = os.path.splitext(os.path.basename(fname))[0]
    return name.replace("-", " ").replace("_", " ").title()

def classify_room_priority(name):
    name = name.lower()
    MAIN = ["wohnzimmer", "flur", "schlafzimmer", "kinderzimmer"]
    FUNC = ["küche", "bad", "balkon", "garage", "wc"]
    for i, r in enumerate(MAIN):
        if r in name:
            return (0, i)
    for i, r in enumerate(FUNC):
        if r in name:
            return (1, i)
    return (2, 99)

def extract_number(fname):
    match = re.search(r'(\d+)(?=\.\w+$)', fname)
    return int(match.group(1)) if match else 999

def add_image_gallery_from_folder(doc, folder, title="Innenansichten", per_row=2, width=2.5):
    if not os.path.isdir(folder):
        return
    images = [os.path.join(folder, f) for f in os.listdir(folder)
              if f.lower().endswith((".png", ".jpg", ".jpeg"))]
    images.sort(key=lambda p: (classify_room_priority(p), extract_number(p)))
    if not images:
        return

    doc.add_heading(title, level=2)
    for i in range(0, len(images), per_row):
        table = doc.add_table(rows=1, cols=per_row)
        row = table.rows[0].cells
        for j in range(per_row):
            if i + j < len(images):
                img = images[i + j]
                p = row[j].paragraphs[0]
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                p.add_run().add_picture(img, width=Inches(width))
                row[j].add_paragraph(clean_filename_for_caption(img)).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# ---------- Save Word File ----------

def save_to_docx_with_images(text, logo_path=None, main_image_path=None,
                             detail_image_folder=None, output_dir="output",
                             title_prefix="anzeige"):

    os.makedirs(output_dir, exist_ok=True)
    filename = f"{title_prefix}-{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    filepath = os.path.join(output_dir, filename)

    doc = Document()
    cleaned = clean_markdown(text)

    add_logo_top_right(doc, logo_path)
    doc.add_paragraph()

    start_index = 0

    if not cleaned or cleaned[0][0] != 'heading1':
        h = doc.add_heading("Immobilien-Exposé", level=1)
        h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        # start_index = 1
    else: 
        title_text = ''.join([t[1] for t in cleaned[0][1]])
        h = doc.add_heading(title_text, level=1)
        h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        start_index = 1

    add_main_image(doc, main_image_path, "Außenansicht der Immobilie")

    # table = None
    bullet_stack = [0]
    table_rows = []

    for elem_type, content in cleaned[start_index:]:
        if elem_type == "table_row":
            table_rows.append(content)
            continue
        elif table_rows:
            # when we are done with the table, render the table
             add_clean_table_to_docx(doc, table_rows)
             table_rows = []

        if elem_type == "heading2":
            doc.add_heading(content, level=2)
            bullet_stack = [0]

        elif elem_type == "bullet":
            p = doc.add_paragraph(style='List Bullet')
            for t, val in content[1]:
                run = p.add_run(val)
                run.bold = (t == 'bold')

        elif elem_type in ("bold_text", "paragraph"):
            p = doc.add_paragraph()
            for t, val in content:
                run = p.add_run(val)
                run.bold = (t == 'bold')

    # Tabelle am Ende noch rendern, falls sie zuletzt kommt
    if table_rows:
        add_clean_table_to_docx(doc, table_rows)

    if detail_image_folder:
        add_image_gallery_from_folder(doc, detail_image_folder)

    doc.save(filepath)
    return filepath