import streamlit as st
import os
import re
from utils import (
    clean_markdown,
    classify_room_priority,
    extract_number,
    create_prompt,
    generate_real_estate_ad_deeps,
    save_to_docx_with_images, 
    rescale_img,
    trim_trailing_notes
)

# Page setup
st.set_page_config(page_title="Immobilien-Anzeigengenerator", page_icon="🏠", layout="wide")


def set_sidebar():
    with st.sidebar:
        st.header("Eingaben")
        AUSSTATTUNG_OPTIONS = [
            "Garage/Stellplatz", "Garten", "Balkon", "Terasse", "Einbauküche",
            "Personenaufzug", "Keller", "Gäste-WC", "Smart Home", "Barrierefrei", "Neubau"
        ]
        key_facts = {
            "Objekt-Nr": st.text_input("Objekt-Nr", "12345"),
            "Wohnfläche": st.text_input("Wohnfläche (m²)", "150"),
            "Grundstücksfläche": st.text_input("Grundstücksfläche (m²)", "250"),
            "Maklerprovision": st.text_input("Maklerprovision", "3,57% inkl. MwSt."),
            "Bezugstermin": st.text_input("Bezugstermin", "ab sofort")
        }
        # Property fields first
        property_inputs = {
            "action": st.selectbox("Aktion*", ["Kauf", "Miete"]),
            "immotype": st.selectbox("Immobilientyp*", ["Haus", "Wohnung"]),
            "location": st.text_input("Ort*", "Berlin"),
            "address": st.text_input("Adresse", "Musterstraße 1, 12345 Berlin"),
            "year_built": st.number_input("Baujahr*", 1700, 2025),
            "living_area": key_facts["Wohnfläche"],
            "price": st.text_input("Preis (€)*", "500.000 €"),
            "heating": st.selectbox("Heizungsart", ["Zentralheizung", "Etagenheizung"]),
            "internet_speed": st.selectbox("Internetgeschwindigkeit", [
                "min. 100 MBit/s", "min. 250 MBit/s", "min. 1000 MBit/s - Glasfaser"]),
            "energy_efficiency_class": st.selectbox("Energieeffizienzklasse", ["A+", "A", "B", "C", "D", "E", "F", "G", "H"]),
            "energy_consumption": st.text_input("Energieverbrauch (kWh/m²a)", "100"),
            "rooms": st.number_input("Zimmer", 1, 20, 4),
            "lot_size": key_facts["Grundstücksfläche"],
            "house_type": st.selectbox("Haustyp", ["Einfamilienhaus", "Doppelhaus", "Reihenhaus"]),
            "features": ", ".join(st.multiselect("Ausstattung/Merkmale", AUSSTATTUNG_OPTIONS)),
            "key_facts": key_facts,
            "specific_title": st.text_input("Objekttitel (z.B. 'Einziehen und wohlfühlen - ...')", "Einziehen und wohlfühlen - Ihr neues Zuhause in Berlin!"),
            "call_to_action": st.text_input("Call-to-Action", "Jetzt Besichtigung vereinbaren – kostenfrei und unverbindlich")
        }
        # Contact fields at the bottom
        st.markdown("---")
        st.header("Makler-Kontakt")
        contact_fields = {
            "name": st.text_input("Ansprechpartner/in", "Max Mustermann"),
            "phone": st.text_input("Telefon", "01234/567890"),
            "email": st.text_input("E-Mail", "info@immobilien.de"),
            "website": st.text_input("Webseite", "www.immobilien.de")
        }
        property_inputs["contact"] = f"{contact_fields['name']}\nTelefon: {contact_fields['phone']}\nE-Mail: {contact_fields['email']}\nWeb: {contact_fields['website']}"
        property_inputs["contact_fields"] = contact_fields
        return property_inputs


def initialize_state():
    st.session_state.setdefault("generated_text", "")
    st.session_state.setdefault("docx_config", {})
    st.session_state.setdefault("docx_path", "")


def main():
    initialize_state()
    st.title("🏠 Immobilienanzeigen-Generator")
    user_input = set_sidebar()
    if st.button("Anzeige generieren"):
        try:
            prompt = create_prompt(user_input)
            st.spinner("Generiere Anzeige...")
            api = st.secrets["deepseek"]
            response = generate_real_estate_ad_deeps(api, prompt)
            response_trimmed = trim_trailing_notes(response)
            st.session_state["generated_text"] = response_trimmed
            st.session_state["docx_config"] = {
                "logo_path": "data/logo-company.png",
                "main_image_path": "data/pics/foto-haus-main.png",
                "detail_image_folder": "data/pics/detailansicht",
                "output_dir": "output",
                "title_prefix": "anzeige",
                "contact_info": user_input["contact"],
                "key_facts": user_input["key_facts"],
                "specific_title": user_input["specific_title"],
                "address": user_input["address"],
                "call_to_action": user_input["call_to_action"],
                "contact_fields": user_input["contact_fields"],
            }
            rescaled_main = rescale_img(st.session_state["docx_config"]["main_image_path"], output_folder="data/rescaled_main")
            st.session_state["docx_config"]["main_image_path"] = rescaled_main
            rescaled_gallery = rescale_img(st.session_state["docx_config"]["detail_image_folder"], max_width=550, quality=80)
            if rescaled_gallery:
                st.session_state["docx_config"]["detail_image_folder"] = os.path.dirname(list(rescaled_gallery.values())[0])
        except Exception as e:
            st.error(f"Fehler: {e}")
    if st.session_state["generated_text"]:
        try:
            st.subheader("📄 Vorschau")
            logo = st.session_state["docx_config"]["logo_path"]
            if os.path.isfile(logo):
                st.columns([4, 1])[1].image(logo, width=120)
            cleaned_parts = clean_markdown(st.session_state["generated_text"])
            table_rows = []
            last_elem_type = None
            last_content = None
            key_facts_shown = False
            i = 0
            while i < len(cleaned_parts):
                elem_type, content = cleaned_parts[i]
                # Only show section titles if they have content after them
                if elem_type in ("heading2", "heading3"):
                    # Remove duplicate/empty/irrelevant headers
                    if content.strip().lower() in ("key facts", "bildergalerie", "(platzhalter für bilder)"):
                        # Only show 'Key Facts' once and only if followed by table
                        if content.strip().lower() == "key facts":
                            # Check if next is a table_row with content
                            j = i + 1
                            while j < len(cleaned_parts) and cleaned_parts[j][0] in ("heading2", "heading3"):
                                j += 1
                            if j < len(cleaned_parts) and cleaned_parts[j][0] == "table_row" and not key_facts_shown:
                                st.markdown("### Key Facts")
                                key_facts_shown = True
                        i += 1
                        continue
                    # Only show header if next non-header is content
                    j = i + 1
                    has_content = False
                    while j < len(cleaned_parts):
                        next_elem, next_content = cleaned_parts[j]
                        if next_elem not in ("heading2", "heading3") and (
                            (next_elem == "table_row" and next_content) or (next_elem in ("paragraph", "bullet", "bold_text") and next_content)):
                            has_content = True
                            break
                        elif next_elem not in ("heading2", "heading3"):
                            break
                        j += 1
                    if has_content:
                        if elem_type == "heading2":
                            st.subheader(content)
                        else:
                            st.markdown(f"**{content}**")
                    i += 1
                    continue
                elif elem_type == "table_row":
                    if not key_facts_shown:
                        st.markdown("### Key Facts")
                        key_facts_shown = True
                    table_rows.append(content)
                elif table_rows:
                    for row in table_rows[1:]:
                        if len(row) == 2:
                            st.markdown(f"- {row[0]}: {row[1]}")
                        else:
                            st.markdown(f"- {' - '.join(row)}")
                    table_rows = []
                elif elem_type == "bullet":
                    for t, val in content[1]:
                        if val.strip():
                            bullet = f"- **{val}**" if t == "bold" else f"- {val}"
                            st.markdown(bullet)
                elif elem_type in ("bold_text", "paragraph"):
                    # Remove any placeholder for images
                    line = ""
                    for t, val in content:
                        if "platzhalter für bilder" in val.lower():
                            continue
                        if t == "bold":
                            line += f"**{val}**"
                        else:
                            line += val
                    if line.strip():
                        st.markdown(line)
                last_elem_type = elem_type
                last_content = content
                i += 1
            if table_rows:
                for row in table_rows[1:]:
                    if len(row) == 2:
                        st.markdown(f"- {row[0]}: {row[1]}")
                    else:
                        st.markdown(f"- {' - '.join(row)}")
            folder = st.session_state["docx_config"]["detail_image_folder"]
            if os.path.isdir(folder):
                st.subheader("🖼️ Innenansichten")
                imgs = sorted([
                    os.path.join(folder, f) for f in os.listdir(folder)
                    if f.lower().endswith((".png", ".jpg", ".jpeg"))
                ], key=lambda p: (
                    classify_room_priority(os.path.basename(p)),
                    extract_number(os.path.basename(p)),
                    os.path.basename(p)
                ))
                for i in range(0, len(imgs), 2):
                    cols = st.columns(2)
                    for j in range(2):
                        if i + j < len(imgs):
                            path = imgs[i + j]
                            filename = os.path.basename(path)
                            caption_filename = os.path.splitext(filename)[0].replace("-", " ").replace("_", " ").title()
                            cols[j].image(path, caption=caption_filename)
            if st.button("📥 Word-Datei erstellen"):
                # Build full property facts for docx
                facts_vertical = [
                    ("Objekt-Nr", user_input["key_facts"]["Objekt-Nr"]),
                    ("Aktion", user_input["action"]),
                    ("Immobilientyp", user_input["immotype"]),
                    ("Ort", user_input["location"]),
                    ("Adresse", user_input["address"]),
                    ("Baujahr", user_input["year_built"]),
                    ("Wohnfläche", user_input["living_area"]),
                    ("Preis", user_input["price"]),
                    ("Heizungsart", user_input["heating"]),
                    ("Internetgeschwindigkeit", user_input["internet_speed"]),
                    ("Energieeffizienzklasse", user_input["energy_efficiency_class"]),
                    ("Energieverbrauch", user_input["energy_consumption"]),
                    ("Zimmer", user_input["rooms"]),
                    ("Grundstücksgröße", user_input["lot_size"]),
                    ("Haustyp", user_input["house_type"]),
                    ("Ausstattung/Merkmale", user_input["features"]),
                    ("Maklerprovision", user_input["key_facts"]["Maklerprovision"]),
                    ("Bezugstermin", user_input["key_facts"]["Bezugstermin"]),
                ]
                st.session_state["docx_config"]["facts_vertical"] = facts_vertical
                docx_path = save_to_docx_with_images(
                    text=st.session_state["generated_text"],
                    **st.session_state["docx_config"]
                )
                st.session_state["docx_path"] = docx_path
                with open(docx_path, "rb") as f:
                    st.download_button("⬇️ Word-Datei herunterladen", f, os.path.basename(docx_path),
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Fehler bei der Vorschau: {e}")


if __name__ == "__main__":
    main()
