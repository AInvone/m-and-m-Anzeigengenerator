import streamlit as st
import os 
from datetime import datetime
from pathlib import Path
import yaml
import re

import utils
from utils import classify_room_priority, extract_number

# Initialisiere Streamlit immer ZUERST
st.set_page_config(
    page_title="Immobilien-Anzeigengenerator",
    page_icon="🏠",
    layout="wide"
)

def set_sidebar():
    with st.sidebar:
        st.header("Eingaben")
        
        AUSSTATTUNG_OPTIONS = ["Garage/Stellplatz", "Garten", "Balkon", "Terasse", "Einbauküche", "Personenaufzug", "Keller", "Gäste-WC", "Smart Home", "Barrierefrei", "Neubau"]
        usr_input = {
            # Pflichtfelder
            "action": st.selectbox("Aktion*", ["Kauf", "Miete"], index=0),
            "immotype": st.selectbox("Immobilientyp*", ["Haus", "Wohnung"]),
            "location": st.text_input("Ort*", "Berlin"),
            "year_built": st.number_input("Baujahr*", min_value=1700, max_value=2025),
            "living_area": st.text_input("Wohnfläche (m²)", "150"),
            "price": st.number_input("Preis (€)*", min_value=0, value=500000),
            "heating": st.selectbox("Heizungsart", ["Zentralheizung", "Etagenheizung"], index=0),
            "internet_speed": st.selectbox("Internetgeschwindigkeit", ["min. 100 MBit/s", "min. 250 MBit/s", "min. 1000 MBit/s - Glasfaser"], index=0),
            "energy_efficiency_class": st.selectbox("Energieeffizienzklasse", ["A+", "A", "B", "C", "D", "E", "F", "G", "H"], index=0),
            # Optionale Felder
            "rooms": st.number_input("Zimmer", min_value=1, value=10),
            "lot_size": st.text_input("Grundstücksgröße (m²)", "250"),
            "house_type": st.selectbox("Haustyp", ["Einfamilienhaus", "Doppelhaus", "Reihenhaus"], index=0),
            "features": ", ".join(
                st.multiselect(
                    label="Ausstattung/Merkmale (Mehrfachauswahl)", 
                    options=AUSSTATTUNG_OPTIONS, 
                    help="Wählen Sie alle zutreffenden Ausstattungsmerkmale aus")
                    )
        }
    return usr_input



def main():
    st.title("🏠 Immobilienanzeigen-Generator")
    
    # Sidebar für Eingaben
    usr_input = set_sidebar()

    # Hauptbereich
    if st.button("Anzeige generieren", type="primary"):
        try:
            # Parameter sammeln
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

            config_file_name = "config.yaml"
            cur_script_folder = os.path.dirname(os.path.abspath(__file__))
            config_file_path = os.path.join(cur_script_folder, '../configs/', config_file_name)

            with open(config_file_path, "r", encoding="utf-8") as f:
                args = yaml.safe_load(f)
            print("CONFIG-all: ", args)
            print("CONFIG deepseek: ", args['deepseek'])
            print("CONFIG openai: ", args['openai'])
            print("CONFIG output: ", args['output'])
            input_logo_path = args['input']['logo_path']
            input_main_image_path = args['input']['main_image_path']
            input_detail_image_folder = args['input']['detail_image_folder']
            print("CONFIG input logo: ", input_logo_path)
            print("CONFIG input main_pic: ", input_main_image_path)
            print("CONFIG input detail_pics: ", input_detail_image_folder)

            # Generiere Anzeige
            with st.spinner("Generiere Anzeige..."):
                prompt = utils.create_prompt(params)
                ad_text = utils.generate_real_estate_ad_deeps(args['deepseek'], prompt)

                # Store only text + paths for now, create docx or pdf upon request
                st.session_state["generated_text"] = ad_text
                st.session_state["docx_config"] = {
                    "logo_path": input_logo_path,
                    "main_image_path": input_main_image_path,
                    "detail_image_folder": input_detail_image_folder,
                    "output_dir": args["output"]["output_directory"],
                    "title_prefix": "anzeige"
                }

        
        except Exception as e:
            st.error(f"Fehler: {str(e)}")         
            
    if "generated_text" in st.session_state:
        try:
            # Ergebnisse anzeigen
            st.subheader("Vorschau")

            input_logo_path = st.session_state["docx_config"]['logo_path']

            # 1. Logo in top-right
            logo_cols = st.columns([4, 1])
            with logo_cols[1]:
                st.image(input_logo_path, width=150)
            
            # Add some space between logo and content. 
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 2. Extract and show title (first line from ad_text) 
            title_line = st.session_state["generated_text"].strip().split("\n")[0]

            # Use regex to remove a pair of asterisks at the start and end
            match = re.match(r"^\*\*(.+?)\*\*$", title_line.strip())
            if match:
                title_line = match.group(1)

            st.markdown(f"<h2 style='text-align: center'>{title_line}</h2>", unsafe_allow_html=True)


            # 3. Main image, centered
            input_main_image_path = st.session_state["docx_config"]['main_image_path']
            cols = st.columns([1, 2, 1])
            with cols[1]:
                st.image(input_main_image_path, caption="Außenansicht der Immobilie", use_container_width=True)
                
            # 4. Rest of exposé text
            rest_of_text = "\n".join(st.session_state["generated_text"].strip().split("\n")[1:]).strip()
            st.markdown(rest_of_text)

            # Innenansichten
            input_detail_image_folder = st.session_state["docx_config"]['detail_image_folder']

            if os.path.isdir(input_detail_image_folder):
                st.subheader("🖼️ Innenansichten")
                image_files = [
                    os.path.join(input_detail_image_folder, f)
                    for f in os.listdir(input_detail_image_folder)
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
                            image_path = sorted_images[i + j]
                            caption = os.path.splitext(os.path.basename(image_path))[0].replace("-", " ").capitalize()
                            cols[j].image(image_path, caption=caption, use_container_width=True)

            # Download-approach 2

            if st.button("📥 Word-Datei erstellen"):
                # Save the docx now
                docx_path = utils.save_to_docx_with_images(
                    text=st.session_state["generated_text"],
                    **st.session_state["docx_config"]
                )
                st.session_state["docx_path"] = docx_path
                st.success(f"Word-Datei gespeichert unter:\n`{os.path.abspath(docx_path)}`")

        except Exception as e:
            st.error(f"Fehler: {str(e)}")


# WICHTIG: Main-Aufruf nicht vergessen!
if __name__ == "__main__":
    main() 
