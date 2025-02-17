# data_logic.py
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

def find_table_starting_from_columns(excel_bytes, sheet_name, required_columns):
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

def create_aggregated_data(filtered_df):
    aggregated = filtered_df.groupby("Naam").size().reset_index(name="Aantal Taken")
    return aggregated

def compare_tasks_grouped_by_name(current_df, previous_df):
    # Zoek de kolom voor Oblig o extern formatteren
    ob_col = next((col for col in current_df.columns if "Obligo extern formaa" in col), None)
    if not ob_col:
        ob_col = "OH-order"
    all_names = pd.unique(pd.concat([current_df["Naam"], previous_df["Naam"]]))
    comparison_data = []
    for name in all_names:
        current_tasks = current_df[current_df["Naam"] == name]
        prev_tasks = previous_df[previous_df["Naam"] == name]
        current_set = set(current_tasks[ob_col].dropna().astype(str))
        prev_set = set(prev_tasks[ob_col].dropna().astype(str))
        common = current_set.intersection(prev_set)
        new_tasks = len(current_set - prev_set)
        old_tasks_count = len(common)
        common_str = ", ".join(common) if common else ""
        comparison_data.append({
            "Naam": name,
            "Aantal nieuwe taken": new_tasks,
            "Aantal oude taken": old_tasks_count,
            f"{ob_col} (oude taken)": common_str
        })
    return pd.DataFrame(comparison_data)

def create_combined_excel_file(
    everything_df=None,
    dict_per_name=None,
    aggregated_df=None,
    comparison_df=None,
    download_everything=False,
    download_per_name=False,
    download_aggregated=False,
    download_comparison=False
):
    """
    Schrijft de beschikbare DataFrames naar één Excel-bestand met meerdere bladen.
    Iedere output komt op een apart werkblad met een toepasselijke naam, zoals:
    'AllesBijElkaar', per persoon (bijvoorbeeld de naam van de persoon), 'GegroepeerdOverzicht' en 'Vergelijking'.
    """
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book
    # Formatter voor datumcellen (short date: mm/dd/yyyy)
    date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})

    def write_sheet(df, sheet_name):
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        # Header formatter: standaard blauwe kleur met witte tekst
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1
        })
        # Schrijf de header opnieuw met de blauwe stijl
        for col_num, col in enumerate(df.columns):
            worksheet.write(0, col_num, col, header_format)
        # Stel de kolombreedte in op basis van de data en de header
        for idx, col in enumerate(df.columns):
            series = df[col].astype(str)
            max_len = max(series.map(len).max(), len(str(col))) + 2
            worksheet.set_column(idx, idx, max_len)
        # Bepaal de data range (exclusief header)
        num_rows = df.shape[0]  # aantal data-rijen
        num_cols = df.shape[1]
        first_data_row = 1
        last_data_row = num_rows  # want header is in rij 0

        # Definieer formaten voor de afwisselende rijen
        even_format = workbook.add_format({'bg_color': '#DCE6F1'})  # lichte blauwe tint
        odd_format = workbook.add_format({'bg_color': '#B8CCE4'})   # iets donkerder blauw

        # Pas condionele opmaak toe: (let op: Excel's ROW() is 1-indexed)
        worksheet.conditional_format(first_data_row, 0, last_data_row, num_cols - 1, {
            'type': 'formula',
            'criteria': '=MOD(ROW(),2)=0',
            'format': even_format
        })
        worksheet.conditional_format(first_data_row, 0, last_data_row, num_cols - 1, {
            'type': 'formula',
            'criteria': '=MOD(ROW(),2)=1',
            'format': odd_format
        })

        # Pas datumformattering toe voor elke kolom die "Leverdatum" bevat
        date_cols = [i for i, col in enumerate(df.columns) if "Leverdatum" in col]
        for col_idx in date_cols:
            for row_num, value in enumerate(df.iloc[:, col_idx], start=1):
                if pd.notna(value) and isinstance(value, (pd.Timestamp, datetime)):
                    worksheet.write_datetime(row_num, col_idx, value, date_format)

    if download_everything and everything_df is not None:
        write_sheet(everything_df, "AllesBijElkaar")
    if download_per_name and dict_per_name is not None:
        for name, df_group in dict_per_name.items():
            # Excel sheetnaam mag maximaal 31 tekens bevatten, dus knippen indien nodig
            valid_name = name if len(name) <= 31 else name[:31]
            write_sheet(df_group, valid_name)
    if download_aggregated and aggregated_df is not None:
        write_sheet(aggregated_df, "GegroepeerdOverzicht")
    if download_comparison and comparison_df is not None:
        write_sheet(comparison_df, "Vergelijking")

    writer.close()
    output.seek(0)
    return output
