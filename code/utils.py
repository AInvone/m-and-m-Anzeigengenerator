import re
from datetime import datetime 
from docx import Document 
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT 

import os 
import yaml 

from openai import OpenAI


def generate_real_estate_ad_gpt4(args, prompt):
    client = OpenAI(
        api_key=args['api_key'],  
        base_url=args['base_url'],
    )
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message['content']


def generate_real_estate_ad_deeps(args, prompt):
    client = OpenAI(
        api_key = args['api_key'], 
        base_url = args['base_url'],
    )
    response = client.chat.completions.create(
        model = args['model'],
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    return response.choices[0].message.content


def create_prompt(params):

    prompt = f"""Erstelle eine professionelle Immobilienanzeige für Immoscout24 auf Deutsch mit diesen Angaben:
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
Energieeffizienzklasse: {params['internet_speed']}
Ausstattung/Merkmale: {params['features']}

Bitte verwende eine übersichtliche Struktur: 
- Eine ansprechende, aufmerksamkeitsstarke Überschrift
- ein kleiner Absatz "Key Facts" mit den vier Hauptmerkmalen: Wohnfläche, Anzahl Zimmer, Grundstück und Preis 
- eine Übersicht mit allen Merkmalen in zwei-spaltiger Auflistung
- Einen ausführlichen Fließtext mit starkem Einstieg und klaren Vorteilen (USPs), mindestens eine halbe DIN A4 Seite lang
- Ein professioneller und verkaufsfördernder Ton
- Zielgruppe: Käufer auf dem deutschen Immobilienmarkt
- Stil: inspirierend, klar, realistisch, ohne Übertreibungen
"""
    # - Professioneller Ton mit Immoscout-Keywords

    return prompt 


# helper function for Line-Bold-Formatting     
def _process_markdown_line(line: str) -> list[tuple[str,str]]:
    """Process individual line for markdown cleaning.
    Args:
        line: Input text line with markdown
    Returns:
        List of (format_type, text) tuples
    """
    parts = []
    buffer = []
    in_bold = False 
    for char in line:
        if char == "*" and len(buffer) >= 1 and buffer[-1]=="*":
            buffer.pop()
            if in_bold:
                parts.append(['bold', ''.join(buffer)])
                in_bold = False 
            else:
                parts.append(('text', ''.join(buffer)))
                in_bold = True 
            buffer = []
        else:
            buffer.append(char)
    if buffer:
        parts.append(['text', ''.join(buffer)])
        
    return parts 


def clean_markdown(text):
    """complete removal of Markdown-formatting and prepare for docx"""
    # text = re.sub(r'\*{2,}(.*?)\*{2,}', r'\1', text)  # **Fett** → Fett
    # text = re.sub(r'#{2,}', '', text)  # ## Überschrift → Überschrift
    # text = re.sub(r'[-•✔🔹]', '', text)  # Bullet-Symbole entfernen

    lines = []
    current_section = None 

    for line in text.split("\n"):

        line = line.strip()
        if not line:
            continue 

        # Recognize strong title on first line (starts & ends with ** or likely heading)
        if not lines and line and (
            (line.startswith("**") and line.endswith("**")) or 
            (len(line) > 20 and line.count(" ") > 3)
        ):
            clean_title = re.sub(r"[*_#]", "", line).strip()
            lines.append(("heading1", [("text", clean_title)]))
            continue

        # Identify Title (### or **)
        # Recognize strong title on first line: heading detection – only on the first non-empty line
        if not lines and line and (
            (line.startswith("**") and line.endswith("**")) or 
            (len(line) > 20 and line.count(" ") > 3)
        ):
            clean_title = re.sub(r"[*_#]", "", line).strip()
            lines.append(("heading1", [("text", clean_title)]))
            continue
        elif line.startswith("###"):
            line = line.replace("###", "").strip()
            lines.append(("heading2", line))
            current_bullet_level = 0
            continue 
        elif line.startswith("**"): 
            line = line.replace("**", "").strip()
            lines.append(("heading3", line))
            current_bullet_level = 0
            continue 
        

        # Bullet Points ()
        bullet_match = re.match((r'^(\s*)([-•✔🔹])\s*(.*)'), line)
        if bullet_match: 
            indent, symbol, content = bullet_match.groups()
            level = len(indent) // 2 + 1  # 2 spaces per level 
            lines.append(("bullet", (level, _process_markdown_line(content))))  # Neue Liste
            current_bullet_list = level
            continue 
        
        # normal text with bold formatting 
        processed = _process_markdown_line(line) 
        if any(t[0] == 'bold' for t in processed):
            lines.append(('bold_text', processed))
        else:
            lines.append(('paragraph', processed))      
    
    return lines


def add_logo_top_right(doc, logo_path):
    if logo_path and os.path.exists(logo_path):
        section = doc.sections[0]
        header = section.header
        paragraph = header.paragraphs[0]
        # Clear any existing content
        paragraph.clear()
        # Align the paragraph right
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        # Add logo aligned right
        run = paragraph.add_run()
        run.add_picture(logo_path, width=Inches(2))

        # Add spacing to avoid overlap with body text
        # Increase top margin to make space for the logo in the header
        section.top_margin = Inches(1.2)  # default is ~1.0", increase to create buffer

    else:
        print("Logo file not found:", logo_path)


def add_main_image(doc, image_path, caption=""):
    print("DEBUG — image_path type:", type(image_path), "value:", image_path)
    if image_path and os.path.exists(image_path):
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run()
        run.add_picture(image_path, width=Inches(5.5))
        if caption:
            caption_paragraph = doc.add_paragraph(caption, style="Caption")
            caption_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# Define room category priorities
MAIN_SPACE = ["wohnzimmer", "eingang", "flur", "schlafzimmer", "kinderzimmer"]
FUNCTIONAL_SPACE = ["küche", "bad", "badezimmer", "balkon", "terasse", "hof", "garage", "stellplatz", "wc", "toilette"]

def extract_number(filename):
    # Extract trailing number, e.g. wohnzimmer-2.jpg → 2
    match = re.search(r'(\d+)(?=\.\w+$)', filename)
    return int(match.group(1)) if match else float('inf')

def room_index(room_name, reference_list):
    for i, r in enumerate(reference_list):
        if r in room_name:
            return i
    return float('inf')  # put unknowns last

def classify_room_priority(filename):
    lower_name = filename.lower()
    for room in MAIN_SPACE:
        if room in lower_name:
            return (0, room_index(lower_name, MAIN_SPACE))
    for room in FUNCTIONAL_SPACE:
        if room in lower_name:
            return (1, room_index(lower_name, FUNCTIONAL_SPACE))
    return (2, float('inf'))  # Unknown category last


def clean_filename_for_caption(filename):
    base = os.path.basename(filename)
    name = os.path.splitext(base)[0]
    name = name.replace("-", " ").replace("_", " ")
    return name.capitalize()


def add_image_gallery_from_folder(doc, folder_path, title="Detailansichten", images_per_row=2, image_width=2.5):
    if not os.path.isdir(folder_path):
        print(f"Ordner nicht gefunden: {folder_path}")
        return

    image_files = []
    for f in os.listdir(folder_path):
        if isinstance(f, tuple):
            f = f[0]  # unwrap tuple
        if isinstance(f, str) and f.lower().endswith((".png", ".jpg", ".jpeg")):
            full_path = os.path.join(folder_path, f)
            if os.path.isfile(full_path):
                image_files.append(full_path)

    print("DEBUG — filtered image_files:")
    for f in image_files:
        print("   ", f, type(f))
    
    if not image_files:
        print("Keine Bilder im angegebenen Ordner gefunden.")
        return

    print("DEBUG — raw files in folder:")
    for f in os.listdir(folder_path):
        print("   ", f, type(f))
        
    sorted_images = sorted(
        image_files,
        key=lambda path: (
            classify_room_priority(os.path.basename(path)),
            extract_number(os.path.basename(path)),
            os.path.basename(path).lower()
        )
    )

    doc.add_heading(title, level=2)

    for i in range(0, len(sorted_images), images_per_row):
        table = doc.add_table(rows=1, cols=images_per_row)
        row_cells = table.rows[0].cells

        for j in range(images_per_row):
            if i + j < len(sorted_images):
                img_path = sorted_images[i + j]
                cell = row_cells[j]
                p = cell.paragraphs[0]
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                p.add_run().add_picture(img_path, width=Inches(image_width))
                
                # Caption
                caption_text = clean_filename_for_caption(img_path)
                cell.add_paragraph(caption_text).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


def save_to_docx_with_images(text, 
                             logo_path = None, 
                             main_image_path=None,
                             detail_image_folder=None, 
                             output_dir = "output", 
                             title_prefix="anzeige"):

    # Create output-folder if it does not yet exist 
    os.makedirs(output_dir, exist_ok = True)

    # Generate filename with current date+time
    timestamp = datetime.now().strftime("%Y-%m-%d_%H_%M_%S.%f")
    filename = f"{title_prefix}-{timestamp}.docx"    
    
    filepath = os.path.join(output_dir, filename)

    # Create docx document with formatting 
    doc = Document()
    cleaned = clean_markdown(text)

    print(cleaned[0][1])

    # Add logo (top of document)
    add_logo_top_right(doc, logo_path)
    doc.add_paragraph()  # Space below logo

    # Title formatting 
    if cleaned and cleaned[0][0] in ['heading1', 'paragraph']:
        print("ADDING TITLE!")
        title_content = cleaned[0][1]
        
        if isinstance(title_content, str):
            title_text = title_content
        else:
            title_text = ''.join([p[1] for p in title_content if p[0] in ['text', 'bold']])
        
        # Apply title-casing (optional: only capitalize first letter of each word)
        title_text = title_text.title()

        heading = doc.add_heading(title_text, level=1)
        heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        start_index = 1
    else:
        # fallback default title
        heading = doc.add_heading("Immobilien-Exposé", level=1)
        heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        start_index = 0



    bullet_stack = [0]  # for nested lists 

    # Add main image
    add_main_image(doc, main_image_path, caption="Außenansicht der Immobilie")

    # main part formatting 
    for elem in cleaned[start_index:]:
        elem_type, content = elem 
        
        if elem_type == "heading2":
            doc.add_heading(content, level=2)
            bullet_stack = [0]

        elif elem_type == "bullet":
            level, parts = content 
            while level > len(bullet_stack):
                # new list level
                p = doc.add_paragraph()
                p.style = 'List Bullet'
                bullet_stack.append(level)
            
            # current list level
            p = doc.add_paragraph(style = 'List Bullet')
            for part_type, text in parts: 
                run = p.add_run(text)
                run.bold = (part_type == 'bold') 

        elif elem_type == "bold_text":
            p = doc.add_paragraph()
            for part_type, text in content:
                run = p.add_run(text)
                run.bold = (part_type == "bold")

        elif elem_type == "paragraph": 
            p = doc.add_paragraph()
            for part_type, text in content: 
                run = p.add_run(text)
                run.bold = (part_type == 'bold')

        # Listenebenen zurücksetzen bei Absätzen
        if elem_type not in ['bullet'] and bullet_stack != [0]:
            bullet_stack = [0]
            
    # Add gallery of interior pictures: load all interior images from folder, sorted
    if detail_image_folder:
        add_image_gallery_from_folder(doc, detail_image_folder)


    # save file 
    doc.save(filepath)
    print(f"Dokument gespeichert als: {filepath}")
    return filepath


def main():
    # Load config with api_keys etc
        # Load the YAML file
    config_file_name = "config.yaml"
    cur_script_folder = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.join(cur_script_folder, '../configs/', config_file_name)
#    print("CONFIG FILE PATH: ", config_file_path)

    with open(config_file_path, "r", encoding="utf-8") as f:
        args = yaml.safe_load(f)
    print("CONFIG: ", args)
    print("CONFIG: ", args['deepseek'])

    usr_input = {}
    usr_input['action'] = "Kauf"
    usr_input['immotype'] = "Haus"
    usr_input['location'] = "Berlin"
    usr_input['price'] = "500,000€"
    usr_input['rooms'] = "4"
    usr_input['lot_size'] = "250 m²"
    usr_input['year_built'] = "2000"
    usr_input['house_type'] = "Einfamilienhaus"
    usr_input['features'] = "Garage, Garten, Balkon"
    usr_input['heating'] = "Zentralheizung"
    usr_input['internet_speed'] = "min. 250 MBit/s"

    prompt = create_prompt(usr_input)
#    output = generate_real_estate_ad_gpt4(prompt)
    print(prompt)
#    output = generate_real_estate_ad_deeps(args['deepseek'], prompt)
    output = """**Traumhaftes Einfamilienhaus in Berlin – Modern, geräumig & voll ausgestattet**  

Sie suchen ein zeitgemäßes Zuhause mit viel Platz und erstklassiger Ausstattung? Dieses charmante **Einfamilienhaus** in Berlin vereint Komfort, Modernität und eine erstklassige Lage. Mit **4 Zimmern**, einem großzügigen **Grundstück von 250 m²** und hochwertigen Details bietet es perfekte Voraussetzungen für Familien oder anspruchsvolle Käufer.  

### **Ihre Vorteile auf einen Blick:**  
✔ **Moderne Bauweise** – Erbaut im Jahr **2000**, mit zeitloser Architektur und stabiler Bausubstanz  
✔ **Hochwertige Ausstattung** – Inkl. **Garage**, gepflegtem **Garten** und gemütlichem **Balkon**  
✔ **Schnelles Internet** – **Mind. 250 MBit/s** für Homeoffice & Entertainment  
✔ **Effiziente Zentralheizung** – Angenehmes Raumklima & niedrige Energiekosten  
✔ **Perfekte Größe** – Großzügige **4 Zimmer** und viel Freiraum für individuelle Gestaltung  

### **Ausstattungs-Highlights:**  
- **Garage** für sicheres Parken  
- **Garten** mit viel Potenzial zur individuellen Nutzung  
- **Balkon** für entspannte Stunden im Freien  
- **Moderne Heiztechnik** (Zentralheizung)  
- **Zukunftssichere Internetanbindung** (min. 250 MBit/s)  

Dieses Haus bietet nicht nur eine **wertstabile Immobilie**, sondern auch ein **lebenswertes Zuhause** in einer attraktiven Lage. Ob als Kapitalanlage oder Eigenheim – hier stimmt einfach alles.  

**Kontaktieren Sie uns jetzt für eine Besichtigung!**  
🔹 **Preis:** 500.000 €  
🔹 **Standort:** Berlin  
🔹 **Verfügbar:** Zum Kauf  

*Nutzen Sie die Chance – vereinbaren Sie noch heute einen Besichtigungstermin!*  

*(Exposé & weitere Details auf Anfrage)*"""

    print(output)
    output_file = save_to_docx(output, output_dir = args["paths"]["output_directory"], title_prefix="anzeige")
    print(output_file)


if __name__ == '__main__':
    main()