# data_logic.py

import pandas as pd
import re
from io import BytesIO
import zipfile
from datetime import datetime

def find_table_starting_from_columns(excel_bytes, sheet_name, required_columns):
    """
    Doorzoekt de aangewezen Excel-sheet op een rij die alle vereiste kolommen bevat.
    Deze rij stelt vervolgens de header vast voor de DataFrame.
    """
    sheet_data = pd.read_excel(excel_bytes, sheet_name=sheet_name, header=None)
    for row_index in range(sheet_data.shape[0]):
        if all(col in sheet_data.iloc[row_index, :].values for col in required_columns):
            table = pd.read_excel(
                excel_bytes,
                sheet_name=sheet_name,
                header=row_index
            )
            return table
    return None

def apply_filters(df, apply_w_filter):
    """
    Past de filters toe:
    - Verantw. Werkplek moet 'VKS' bevatten
    - Status in ['VRIJ', 'OPEN']
    - Leverdatum <= vandaag
    - (optioneel) 'Omschrijving middel' niet eindigend op getal+'w'
    """
    df["Leverdatum"] = pd.to_datetime(df["Leverdatum"], errors="coerce")
    today = pd.Timestamp("today")

    filtered_df = df[
        df["Verantw. Werkplek"].str.contains("VKS", na=False) &
        df["Status"].isin(["VRIJ", "OPEN"]) &
        (df["Leverdatum"] <= today)
    ]
    if apply_w_filter:
        mask = ~filtered_df["Omschrijving middel"].str.contains(r"\d+w$", na=False)
        filtered_df = filtered_df[mask]

    return filtered_df

def process_filtered_data(df, selected_cols, per_naam=True):
    """
    Selecteert kolommen uit 'selected_cols' en splitst (optioneel) op 'Naam'.
    Retourneert (combined_df, dict_of_groups) als per_naam=True, anders (df, None).
    """
    df_filtered = df[selected_cols].copy()
    if per_naam:
        grouped_dfs = {}
        combined_df = pd.DataFrame(columns=selected_cols)
        for name, group in df_filtered.groupby("Naam"):
            grouped_dfs[name] = group
            combined_df = pd.concat([combined_df, group], ignore_index=True)
        return combined_df, grouped_dfs
    else:
        return df_filtered, None

def create_excel_file(df):
    """
    Maakt een Excel-bestand (BytesIO) zonder alle gekleurde randen of stijlen.
    """
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    df.reset_index(drop=True, inplace=True)
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    writer.close()
    output.seek(0)
    return output

def create_aggregated_excel(df):
    """
    Maakt een Excel-bestand (als BytesIO) met een overzicht per 'Naam'
    en het aantal taken (rijen) dat bij elke naam hoort.
    """
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")

    grouped = df.groupby("Naam").size().reset_index(name="AantalTaken")
    grouped.to_excel(writer, index=False, sheet_name="TakenPerNaam")

    writer.close()
    output.seek(0)
    return output

def create_custom_zip(
    everything_df=None,
    dict_per_name=None,
    aggregated_df=None,
    download_everything=False,
    download_per_name=False,
    download_aggregated=False
):
    """
    Maakt een zip-bestand (BytesIO) met (optioneel):
     - Alles bij elkaar (zonder kleuren)
     - Per naam (losse Excel per unieke naam)
     - Telling per naam (TakenPerNaam.xlsx)
    """
    output_zip = BytesIO()
    with zipfile.ZipFile(output_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as z:
        if download_everything and everything_df is not None:
            file_bytes = create_excel_file(everything_df).getvalue()
            z.writestr("Alle_Taken.xlsx", file_bytes)

        if download_per_name and dict_per_name is not None:
            for name, df_group in dict_per_name.items():
                excel_bytes = create_excel_file(df_group).getvalue()
                valid_name = re.sub(r'[\\/*?:"<>|]', "_", str(name))
                z.writestr(f"{valid_name}.xlsx", excel_bytes)

        if download_aggregated and aggregated_df is not None:
            agg_bytes = create_aggregated_excel(aggregated_df).getvalue()
            z.writestr("TakenPerNaam.xlsx", agg_bytes)

    output_zip.seek(0)
    return output_zip
