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
        contact_fields = {
            "name": st.text_input("Ansprechpartner/in", "Max Mustermann"),
            "phone": st.text_input("Telefon", "01234/567890"),
            "email": st.text_input("E-Mail", "info@immobilien.de"),
            "website": st.text_input("Webseite", "www.immobilien.de")
        }
        return {
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
            "contact": f"{contact_fields['name']}\nTelefon: {contact_fields['phone']}\nE-Mail: {contact_fields['email']}\nWeb: {contact_fields['website']}",
            "specific_title": st.text_input("Objekttitel (z.B. 'Einziehen und wohlfühlen - ...')", "Einziehen und wohlfühlen - Ihr neues Zuhause in Berlin!"),
            "key_facts": key_facts,
            "call_to_action": st.text_input("Call-to-Action", "Jetzt Besichtigung vereinbaren – kostenfrei und unverbindlich"),
            "contact_fields": contact_fields,
            "floorplan_path": st.text_input("Pfad zum Grundriss (optional)", "")
        }


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
                "floorplan_path": user_input["floorplan_path"]
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
            st.info("Hinweis: Beim Klick auf 'Word-Datei herunterladen' öffnet Ihr Browser einen Speichern-Dialog. Sie können den Speicherort frei wählen.")
            logo = st.session_state["docx_config"]["logo_path"]
            if os.path.isfile(logo):
                st.columns([4, 1])[1].image(logo, width=120)
            cleaned_parts = clean_markdown(st.session_state["generated_text"])
            table_rows = []
            last_elem_type = None
            for elem_type, content in cleaned_parts:
                if elem_type == "heading1":
                    st.markdown(f"<h2 style='text-align: center'>{content[0][1]}</h2>", unsafe_allow_html=True)
                elif elem_type == "heading2":
                    st.subheader(content)
                elif elem_type == "heading3":
                    st.markdown(f"**{content}**")
                elif elem_type == "table_row":
                    table_rows.append(content)
                elif table_rows:
                    if last_elem_type not in ("heading2", "heading3"):
                        st.markdown("### Merkmale und Details")
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
                    line = ""
                    for t, val in content:
                        if t == "bold":
                            line += f"**{val}**"
                        else:
                            line += val
                    st.markdown(line)
                last_elem_type = elem_type
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
