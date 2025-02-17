# app.py
import streamlit as st
import pandas as pd
from datetime import datetime
from data_logic import (
    find_table_starting_from_columns,
    apply_filters,
    process_filtered_data,
    create_aggregated_data,
    compare_tasks_grouped_by_name,
    create_combined_excel_file
)

def main():
    st.title("Excel Obligo rapportage")
    st.write(
        """
        1. Upload een Excel-bestand (.xlsx of .xls).  
        2. Selecteer kolommen voor de output (voor 'Alles bij elkaar' en 'Per persoon').  
        3. Filter de automaat orders (ja/nee).  
        4. Kies welke outputs je wilt genereren:  
           - Alles bij elkaar  
           - Per persoon  
           - Gegroepeerd overzicht (naam en aantal taken)  
           - Vergelijking met vorige week
        """
    )

    # Upload huidig weekbestand
    uploaded_file = st.file_uploader("Upload je Excel-bestand", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            # Automatisch het blad selecteren met "DOWNLOAD" (case-insensitive)
            download_sheet = next((s for s in sheets if "DOWNLOAD" in s.upper()), None)
            if download_sheet:
                st.write("Automatisch geselecteerd blad:", download_sheet)
                chosen_sheet = download_sheet
            else:
                chosen_sheet = st.selectbox("Selecteer een blad (actueel):", sheets)

            # Vereiste kolommen voor het vinden van de tabel in het bestand
            required_columns = [
                "OH-planningsgroep",
                "Naam",
                "Status",
                "Omschrijving middel",
                "Verantw. Werkplek",
                "Leverdatum",
                "OH-order"
            ]

            # Haal de tabel op gebaseerd op de vereiste kolommen
            df = find_table_starting_from_columns(
                excel_bytes=uploaded_file,
                sheet_name=chosen_sheet,
                required_columns=required_columns
            )
            if df is None:
                st.error("Geen tabel gevonden met de vereiste kolommen.")
                return

            # Definieer gewenste kolom-substrings
            desired_substrings = [
                "Naam", 
                "OH-order", 
                "Status", 
                "Ord.srt", 
                "Verpl. Srt", 
                "Obligo extern formaa", 
                "Omschrijving middel", 
                "Leverdatum", 
                "Leverancier", 
                "Met SES", 
                "SES ontvangen", 
                "Obligo’s EUR"
            ]

            st.subheader("Selecteer de kolommen voor de outputs 'Alles bij elkaar' en 'Per persoon':")
            selected_cols = []
            for col in df.columns:
                # Vink standaard aan indien de kolomnaam een gewenste substring bevat
                default_checked = any(sub in col for sub in desired_substrings)
                if st.checkbox(col, value=default_checked):
                    selected_cols.append(col)
            if not selected_cols:
                st.warning("Je hebt geen kolommen geselecteerd. Selecteer er minimaal één.")
                return

            # Radio-knop: standaard op 'Ja' voor het filteren van automaat orders
            apply_w_filter = st.radio(
                "Wil je de automaat orders eruit filteren?",
                options=["Nee", "Ja"],
                index=1
            ) == "Ja"
            filtered_df = apply_filters(df, apply_w_filter)

            # Verwerk de outputs 'Alles bij elkaar' en 'Per persoon'
            combined_df, groups_dict = process_filtered_data(
                filtered_df,
                selected_cols,
                per_naam=True
            )

            # Creëer een overzicht met naam en aantal taken
            aggregated_df = create_aggregated_data(filtered_df)

            st.subheader("Kies welke outputs je wilt opnemen")
            download_everything = st.checkbox("Alles bij elkaar", value=True)
            download_per_name = st.checkbox("Per persoon", value=False)
            download_aggregated = st.checkbox("Gegroepeerd overzicht (naam en aantal taken)", value=True)

            # Vergelijking met vorige week
            compare_files = st.checkbox("Vergelijk met vorige week")
            if compare_files:
                previous_file = st.file_uploader("Upload Excel-bestand van vorige week", type=["xlsx", "xls"])
                if previous_file is not None:
                    prev_xls = pd.ExcelFile(previous_file)
                    prev_sheets = prev_xls.sheet_names
                    prev_download_sheet = next((s for s in prev_sheets if "DOWNLOAD" in s.upper()), None)
                    if prev_download_sheet:
                        st.write("Automatisch geselecteerd blad (vorige week):", prev_download_sheet)
                        prev_chosen_sheet = prev_download_sheet
                    else:
                        prev_chosen_sheet = st.selectbox("Selecteer een blad (vorige week):", prev_sheets)
                    prev_df = find_table_starting_from_columns(
                        excel_bytes=previous_file,
                        sheet_name=prev_chosen_sheet,
                        required_columns=required_columns
                    )
                    if prev_df is None:
                        st.error("Geen tabel gevonden in het bestand van vorige week.")
                        return
                    prev_filtered_df = apply_filters(prev_df, apply_w_filter)
                    comparison_df = compare_tasks_grouped_by_name(filtered_df, prev_filtered_df)
                    st.success("Vergelijking gemaakt!")
                    download_comparison = True
                else:
                    st.warning("Upload een bestand van vorige week om de vergelijking te maken.")
                    comparison_df = None
                    download_comparison = False
            else:
                comparison_df = None
                download_comparison = False

            if not (download_everything or download_per_name or download_aggregated or download_comparison):
                st.warning("Selecteer ten minste één output.")
                return

            # Maak één gecombineerde knop voor genereren & downloaden
            date_str = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"output_{date_str}.xlsx"
            combined_excel = create_combined_excel_file(
                everything_df=combined_df if download_everything else None,
                dict_per_name=groups_dict if download_per_name else None,
                aggregated_df=aggregated_df if download_aggregated else None,
                comparison_df=comparison_df if download_comparison else None,
                download_everything=download_everything,
                download_per_name=download_per_name,
                download_aggregated=download_aggregated,
                download_comparison=download_comparison
            )

            st.download_button(
                label="Genereer en download Excel-bestand",
                data=combined_excel.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Er trad een fout op: {e}")

if __name__ == "__main__":
    main()
