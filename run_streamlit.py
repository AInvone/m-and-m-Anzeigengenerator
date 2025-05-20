import streamlit as st
import os
import re
from datetime import datetime
from utils import classify_room_priority, extract_number, create_prompt, generate_real_estate_ad_deeps, save_to_docx_with_images

# Streamlit page settings
st.set_page_config(
    page_title="Immobilien-Anzeigengenerator",
    page_icon="🏠",
    layout="wide"
)

# Sidebar form for user input
def set_sidebar():
    with st.sidebar:
        st.header("Eingaben")

        AUSSTATTUNG_OPTIONS = [
            "Garage/Stellplatz", "Garten", "Balkon", "Terasse", "Einbauküche",
            "Personenaufzug", "Keller", "Gäste-WC", "Smart Home", "Barrierefrei", "Neubau"
        ]

        usr_input = {
            "action": st.selectbox("Aktion*", ["Kauf", "Miete"]),
            "immotype": st.selectbox("Immobilientyp*", ["Haus", "Wohnung"]),
            "location": st.text_input("Ort*", "Berlin"),
            "year_built": st.number_input("Baujahr*", min_value=1700, max_value=2025),
            "living_area": st.text_input("Wohnfläche (m²)", "150"),
            "price": st.number_input("Preis (€)*", min_value=0, value=500000),
            "heating": st.selectbox("Heizungsart", ["Zentralheizung", "Etagenheizung"]),
            "internet_speed": st.selectbox("Internetgeschwindigkeit", ["min. 100 MBit/s", "min. 250 MBit/s", "min. 1000 MBit/s - Glasfaser"]),
            "energy_efficiency_class": st.selectbox("Energieeffizienzklasse", ["A+", "A", "B", "C", "D", "E", "F", "G", "H"]),
            "rooms": st.number_input("Zimmer", min_value=1, value=10),
            "lot_size": st.text_input("Grundstücksgröße (m²)", "250"),
            "house_type": st.selectbox("Haustyp", ["Einfamilienhaus", "Doppelhaus", "Reihenhaus"]),
            "features": ", ".join(
                st.multiselect("Ausstattung/Merkmale", options=AUSSTATTUNG_OPTIONS)
            )
        }
    return usr_input

# Initialize session_state defaults
def initialize_session_state():
    if "generated_text" not in st.session_state:
        st.session_state["generated_text"] = ""
    if "docx_config" not in st.session_state:
        st.session_state["docx_config"] = {}
    if "docx_path" not in st.session_state:
        st.session_state["docx_path"] = ""

# Main application logic
def main():
    initialize_session_state()
    st.title("🏠 Immobilienanzeigen-Generator")
    usr_input = set_sidebar()

    if st.button("Anzeige generieren", type="primary"):
        try:
            params = {
                "action": usr_input['action'],
                "immotype": usr_input['immotype'],
                "location": usr_input['location'],
                "year_built": usr_input['year_built'],
                "living_area": usr_input['living_area'],
                "price": f"{usr_input['price']}€",
                "heating": usr_input['heating'],
                "internet_speed": usr_input['internet_speed'],
                "energy_efficiency_class": usr_input['energy_efficiency_class'],
                "rooms": usr_input['rooms'],
                "lot_size": usr_input['lot_size'],
                "house_type": usr_input['house_type'],
                "features": usr_input['features']
            }

            # Load paths + secrets
            input_logo_path = os.path.join("data", "logo-company.jpg")
            input_main_image_path = os.path.join("data", "pics", "foto-haus-main.jpg")
            input_detail_image_folder = os.path.join("data", "pics", "detailansicht")
            output_dir = "output"

            # API config
            args = {
                "deepseek": {
                    "api_key": st.secrets["deepseek_api_key"],
                    "base_url": st.secrets["deepseek_base_url"],
                    "model": st.secrets["deepseek_model"]
                }
            }

            # Prompt + generation
            with st.spinner("Generiere Anzeige..."):
                prompt = create_prompt(params)
                ad_text = generate_real_estate_ad_deeps(args['deepseek'], prompt)

            st.session_state["generated_text"] = ad_text
            st.session_state["docx_config"] = {
                "logo_path": input_logo_path,
                "main_image_path": input_main_image_path,
                "detail_image_folder": input_detail_image_folder,
                "output_dir": output_dir,
                "title_prefix": "anzeige"
            }

        except Exception as e:
            st.error(f"Fehler bei der Generierung: {str(e)}")

    if st.session_state["generated_text"]:
        try:
            st.subheader("📄 Vorschau")

            # 1. Logo
            logo_path = st.session_state["docx_config"].get("logo_path", "")
            if os.path.exists(logo_path):
                st.columns([4, 1])[1].image(logo_path, width=150)
                st.markdown("<br>", unsafe_allow_html=True)

            # 2. Title
            title_line = st.session_state["generated_text"].strip().split("\n")[0]
            title_line = re.sub(r"^\*\*(.+?)\*\*$", r"\1", title_line.strip())
            st.markdown(f"<h2 style='text-align: center'>{title_line}</h2>", unsafe_allow_html=True)

            # 3. Main image
            main_img = st.session_state["docx_config"].get("main_image_path", "")
            if os.path.exists(main_img):
                st.columns([1, 2, 1])[1].image(main_img, caption="Außenansicht der Immobilie", use_container_width=True)

            # 4. Rest of exposé
            text_body = "\n".join(st.session_state["generated_text"].strip().split("\n")[1:]).strip()
            st.markdown(text_body)

            # 5. Gallery
            detail_folder = st.session_state["docx_config"].get("detail_image_folder", "")
            if os.path.isdir(detail_folder):
                st.subheader("🖼️ Innenansichten")
                image_files = [
                    os.path.join(detail_folder, f)
                    for f in os.listdir(detail_folder)
                    if f.lower().endswith((".png", ".jpg", ".jpeg"))
                ]
                sorted_images = sorted(
                    image_files,
                    key=lambda path: (
                        classify_room_priority(os.path.basename(path)),
                        extract_number(os.path.basename(path)),
                        os.path.basename(path).lower()
                    )
                )
                for i in range(0, len(sorted_images), 2):
                    cols = st.columns(2)
                    for j in range(2):
                        if i + j < len(sorted_images):
                            img_path = sorted_images[i + j]
                            caption = os.path.splitext(os.path.basename(img_path))[0].replace("-", " ").capitalize()
                            cols[j].image(img_path, caption=caption, use_container_width=True)

            # 6. Generate and download docx
            if st.button("📥 Word-Datei erstellen"):
                docx_path = save_to_docx_with_images(
                    text=st.session_state["generated_text"],
                    **st.session_state["docx_config"]
                )
                st.session_state["docx_path"] = docx_path

                with open(docx_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Jetzt Word-Datei herunterladen",
                        data=f,
                        file_name=os.path.basename(docx_path),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success(f"✅ Datei gespeichert unter:\n`{os.path.abspath(docx_path)}`")

        except Exception as e:
            st.error(f"Fehler bei der Vorschau: {str(e)}")

# Entry point
if __name__ == "__main__":
    main()
    