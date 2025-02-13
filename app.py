# app.py

import streamlit as st
import pandas as pd
from datetime import datetime

from data_logic import (
    find_table_starting_from_columns,
    apply_filters,
    process_filtered_data,
    create_custom_zip
)

def main():
    st.title("Excel Obligo rapportage")

    st.write("""
    1. Upload een Excel-bestand (.xlsx of .xls).  
    2. Selecteer kolommen voor de output.  
    3. Automaat orders filteren (ja/nee)
    4. Kies welke outputs je wilt genereren:  
        - Alles bij elkaar  
        - Per persoon  
        - Gegroepeerd overzicht  
    """)

    uploaded_file = st.file_uploader("Upload je Excel-bestand", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            st.write("Beschikbare sheets:", sheets)

            chosen_sheet = st.selectbox("Kies een sheet:", sheets)

            required_columns = [
                "OH-planningsgroep",
                "Naam",
                "Status",
                "Omschrijving middel",
                "Verantw. Werkplek",
                "Leverdatum"
            ]

            df = find_table_starting_from_columns(
                excel_bytes=uploaded_file,
                sheet_name=chosen_sheet,
                required_columns=required_columns
            )

            if df is None:
                st.error("Geen tabel gevonden met de vereiste kolommen.")
                return

            st.subheader("Selecteer de kolommen voor de uiteindelijke output:")
            selected_cols = []
            for col in df.columns:
                default_checked = (col in required_columns)
                if st.checkbox(col, value=default_checked):
                    selected_cols.append(col)

            if not selected_cols:
                st.warning("Je hebt geen kolommen geselecteerd. Selecteer er minimaal één.")
                return

            apply_w_filter = st.radio(
                "Wil je de automaat orders eruit filteren?",
                ["Nee", "Ja"]
            ) == "Ja"

            filtered_df = apply_filters(df, apply_w_filter)

            # Verwerk data en groepeer eventueel per naam
            combined_df, groups_dict = process_filtered_data(
                filtered_df,
                selected_cols,
                per_naam=True
            )

            st.subheader("Kies welke bestanden je wilt downloaden")
            download_everything = st.checkbox("Alles bij elkaar", value=True)
            download_per_name = st.checkbox("Per persoon")
            download_aggregated = st.checkbox("Gegroepeerd overzicht")

            # Voor het gegroepeerde bestand (TakenPerNaam)
            aggregated_df = combined_df if download_aggregated else None

            # Controleer of er ten minste één output-vinkje aanstaat
            if not (download_everything or download_per_name or download_aggregated):
                st.warning("Je hebt geen output-opties geselecteerd. Selecteer ten minste één optie.")
                return

            if st.button("Genereer ZIP-bestand"):
                date_str = datetime.now().strftime("%Y-%m-%d")

                zip_bytes = create_custom_zip(
                    everything_df=combined_df if download_everything else None,
                    dict_per_name=groups_dict if download_per_name else None,
                    aggregated_df=aggregated_df,
                    download_everything=download_everything,
                    download_per_name=download_per_name,
                    download_aggregated=download_aggregated
                )
                zip_filename = f"output_{date_str}.zip"

                st.download_button(
                    label="Download ZIP",
                    data=zip_bytes.getvalue(),
                    file_name=zip_filename,
                    mime="application/x-zip-compressed"
                )

        except Exception as e:
            st.error(f"Er trad een fout op: {e}")

if __name__ == "__main__":
    main()
