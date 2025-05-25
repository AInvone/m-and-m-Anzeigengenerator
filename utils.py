
import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI
from PIL import Image

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
- Einen ausführlichen, umfangreichen und ansprechenden Haupttext in Fließtext mit Hervorhebung aller Vorteile (USPs)
- Die Fließtext-Absätze dürfen länger und ausgeschweifter sein, gerne rethorisch ausgeschmückt, oder auch mit einem oder zwei inhaltlichen "Hooks", bevor es zu "Ihre Vorteile auf einen Blick:" kommt
- Zielgruppe: Käufer auf dem deutschen Immobilienmarkt
- Stil: inspirierend, klar, realistisch, ansprechend

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
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue

        # Markdown table row
        if "|" in line and re.match(r"^\|.*\|$", line):
            cells = [c.strip("* ").strip() for c in line.strip().strip("|").split("|")]
            lines.append(("table_row", [("bold" if "**" in c else "text", c.strip("*")) for c in cells]))
            continue

        # Headings
        if line.startswith("###"):
            lines.append(("heading2", line.replace("###", "").strip()))
            continue
        elif line.startswith("**") and line.endswith("**"):
            lines.append(("heading3", line.strip("*").strip()))
            continue

        # Bullet points – skip bad lines like "• --"
        bullet_match = re.match(r'^\s*[-•✔🔹]\s+(.+)', line)
        if bullet_match:
            lines.append(("bullet", (1, _process_markdown_line(bullet_match.group(1)))))
            continue

        # Paragraph
        processed = _process_markdown_line(line)
        if any(t[0] == 'bold' for t in processed):
            lines.append(("bold_text", processed))
        else:
            lines.append(("paragraph", processed))

    return lines

# ---------- Word Helper ----------

def add_logo_top_right(doc, logo_path):
    if logo_path and os.path.exists(logo_path):
        section = doc.sections[0]
        section.top_margin = Inches(1.2)
        header = section.header
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        p.add_run().add_picture(logo_path, width=Inches(2))

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

# ---------- Image Rescale ----------

def rescale_img(input_path_or_folder, output_folder="data/rescaled_details", max_width=1000, quality=85):
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
    if cleaned and cleaned[0][0] == 'heading1':
        title_text = ''.join([t[1] for t in cleaned[0][1]])
        h = doc.add_heading(title_text, level=1)
        h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        start_index = 1

    add_main_image(doc, main_image_path, "Außenansicht der Immobilie")

    bullet_stack = [0]
    table = None

    for elem_type, content in cleaned[start_index:]:
        if elem_type == "table_row":
            if table is None:
                table = doc.add_table(rows=0, cols=len(content))
                table.style = 'Table Grid'
            row = table.add_row().cells
            for i, (ftype, val) in enumerate(content):
                run = row[i].paragraphs[0].add_run(val)
                run.bold = (ftype == "bold")
            continue
        else:
            table = None

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

    if detail_image_folder:
        add_image_gallery_from_folder(doc, detail_image_folder)

    doc.save(filepath)
    return filepath
