import streamlit as st
import pandas as pd
from datetime import datetime

from data_logic import (
    find_table_starting_from_columns,
    apply_filters,
    process_filtered_data,
    create_aggregated_data,
    create_custom_zip,
    compare_tasks_grouped_by_name
)


def main():
    st.title("Excel Obligo rapportage")

    st.write("""
    1. Upload een Excel-bestand (.xlsx of .xls).  
    2. Selecteer kolommen voor de output (voor 'Alles bij elkaar' en 'Per persoon').  
    3. Automaat orders filteren (ja/nee).  
    4. Kies welke outputs je wilt genereren:  
        - Alles bij elkaar  
        - Per persoon  
        - Gegroepeerd overzicht (naam en aantal taken)  
        - Vergelijking met vorige week  
    """)

    # Upload huidig weekbestand
    uploaded_file = st.file_uploader("Upload je Excel-bestand", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            st.write("Beschikbare sheets:", sheets)

            chosen_sheet = st.selectbox("Selecteer een blad (actueel):", sheets)

            # Vereiste kolommen
            required_columns = [
                "OH-planningsgroep",
                "Naam",
                "Status",
                "Omschrijving middel",
                "Verantw. Werkplek",
                "Leverdatum",
                "OH-order"  # Nodig om OH-orders van taken te identificeren
            ]

            # Zoek vereiste tabellen in huidig bestand
            df = find_table_starting_from_columns(
                excel_bytes=uploaded_file,
                sheet_name=chosen_sheet,
                required_columns=required_columns
            )

            if df is None:
                st.error("Geen tabel gevonden met de vereiste kolommen.")
                return

            # Kolommen selecteren voor outputs "Alles bij elkaar" en "Per persoon"
            st.subheader("Selecteer de kolommen voor de outputs 'Alles bij elkaar' en 'Per persoon':")
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

            # Verwerk data voor outputs Alles bij elkaar en Per persoon
            combined_df, groups_dict = process_filtered_data(
                filtered_df,
                selected_cols,
                per_naam=True
            )

            # Verwerk data voor Gegroepeerd overzicht
            aggregated_df = create_aggregated_data(filtered_df)

            # Downloadbare opties
            st.subheader("Kies welke bestanden je wilt downloaden")
            download_everything = st.checkbox("Alles bij elkaar", value=True)
            download_per_name = st.checkbox("Per persoon")
            download_aggregated = st.checkbox("Gegroepeerd overzicht (naam en aantal taken)")

            # Taken vergelijken
            compare_files = st.checkbox("Vergelijk met vorige week")

            if compare_files:
                previous_file = st.file_uploader("Upload Excel-bestand van vorige week", type=["xlsx", "xls"])

                if previous_file is not None:
                    # Vraag welk blad geselecteerd moet worden voor vorige week
                    prev_xls = pd.ExcelFile(previous_file)
                    prev_sheets = prev_xls.sheet_names
                    st.write("Beschikbare sheets in vorige week-bestand:", prev_sheets)

                    prev_chosen_sheet = st.selectbox("Selecteer een blad (vorige week):", prev_sheets)

                    # Lees vorige week-data
                    prev_df = find_table_starting_from_columns(
                        excel_bytes=previous_file,
                        sheet_name=prev_chosen_sheet,
                        required_columns=required_columns
                    )

                    if prev_df is None:
                        st.error("Geen tabel gevonden in het bestand van vorige week.")
                        return

                    prev_filtered_df = apply_filters(prev_df, apply_w_filter)

                    # Vergelijk de huidige taken met vorige week, inclusief oude en nieuwe taken per persoon
                    comparison_df = compare_tasks_grouped_by_name(filtered_df, prev_filtered_df)

                    # Vergelijking automatisch opnemen in download
                    st.success("Vergelijking gemaakt! Dit wordt opgenomen in je ZIP-bestand.")
                    download_comparison = True
                else:
                    st.warning("Upload een bestand van vorige week om de vergelijking te maken.")
                    comparison_df = None
                    download_comparison = False
            else:
                comparison_df = None
                download_comparison = False

            # Controleer of er ten minste één download-optie is geselecteerd
            if not (download_everything or download_per_name or download_aggregated or download_comparison):
                st.warning("Je hebt geen output-opties geselecteerd. Selecteer ten minste één optie.")
                return

            if st.button("Genereer ZIP-bestand"):
                date_str = datetime.now().strftime("%Y-%m-%d")
                output_folder_name = f"output_{date_str}"

                zip_bytes = create_custom_zip(
                    everything_df=combined_df if download_everything else None,
                    dict_per_name=groups_dict if download_per_name else None,
                    aggregated_df=aggregated_df if download_aggregated else None,
                    comparison_df=comparison_df if download_comparison else None,
                    download_everything=download_everything,
                    download_per_name=download_per_name,
                    download_aggregated=download_aggregated,
                    download_comparison=download_comparison,
                    output_folder_name=output_folder_name
                )
                zip_filename = f"{output_folder_name}.zip"

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
