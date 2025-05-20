import streamlit as st
import os
import re
from utils import (
    classify_room_priority,
    extract_number,
    create_prompt,
    generate_real_estate_ad_deeps,
    save_to_docx_with_images
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

        return {
            "action": st.selectbox("Aktion*", ["Kauf", "Miete"]),
            "immotype": st.selectbox("Immobilientyp*", ["Haus", "Wohnung"]),
            "location": st.text_input("Ort*", "Berlin"),
            "year_built": st.number_input("Baujahr*", 1700, 2025),
            "living_area": st.text_input("Wohnfläche (m²)", "150"),
            "price": st.number_input("Preis (€)*", 0, 1_000_000, 500000),
            "heating": st.selectbox("Heizungsart", ["Zentralheizung", "Etagenheizung"]),
            "internet_speed": st.selectbox("Internetgeschwindigkeit", [
                "min. 100 MBit/s", "min. 250 MBit/s", "min. 1000 MBit/s - Glasfaser"]),
            "energy_efficiency_class": st.selectbox("Energieeffizienzklasse", ["A+", "A", "B", "C", "D", "E", "F", "G", "H"]),
            "rooms": st.number_input("Zimmer", 1, 20, 4),
            "lot_size": st.text_input("Grundstücksgröße (m²)", "250"),
            "house_type": st.selectbox("Haustyp", ["Einfamilienhaus", "Doppelhaus", "Reihenhaus"]),
            "features": ", ".join(st.multiselect("Ausstattung/Merkmale", AUSSTATTUNG_OPTIONS))
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
            st.session_state["generated_text"] = response

            st.session_state["docx_config"] = {
                "logo_path": "data/logo-company.jpg",
                "main_image_path": "data/pics/foto-haus-main.jpg",
                "detail_image_folder": "data/pics/detailansicht",
                "output_dir": "output",
                "title_prefix": "anzeige"
            }
        except Exception as e:
            st.error(f"Fehler: {e}")

    if st.session_state["generated_text"]:
        try:
            st.subheader("📄 Vorschau")

            # LOGO
            logo = st.session_state["docx_config"]["logo_path"]
            if os.path.isfile(logo):
                st.columns([4, 1])[1].image(logo, width=120)

            # TITEL
            title = st.session_state["generated_text"].split("\n")[0]
            title = re.sub(r"^\*\*(.+?)\*\*$", r"\1", title.strip())
            st.markdown(f"<h2 style='text-align: center'>{title}</h2>", unsafe_allow_html=True)

            # MAIN IMAGE
            main_img = st.session_state["docx_config"]["main_image_path"]
            if os.path.isfile(main_img):
                st.columns([1, 2, 1])[1].image(main_img, caption="Außenansicht", use_container_width=True)

            # BODY TEXT
            body = "\n".join(st.session_state["generated_text"].split("\n")[1:])
            st.markdown(body)

            # DETAILBILDER
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
                            cols[j].image(path, caption=os.path.basename(path).replace("-", " ").title())

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
    