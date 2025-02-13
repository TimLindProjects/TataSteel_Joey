import pandas as pd
import re
from io import BytesIO
import zipfile
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
    """
    Creëert een overzicht met 'Naam' en 'Aantal taken'.
    """
    aggregated = filtered_df.groupby("Naam").size().reset_index(name="Aantal Taken")
    return aggregated


def compare_tasks_grouped_by_name(current_df, previous_df):
    """
    Vergelijk taken per persoon en geef per naam:
    - Aantal nieuwe taken
    - Aantal oude taken
    - OH-orders van oude taken (gecombineerd als string)
    """
    all_names = pd.unique(pd.concat([current_df["Naam"], previous_df["Naam"]]))
    comparison_data = []

    for name in all_names:
        current_tasks = current_df[current_df["Naam"] == name]
        prev_tasks = previous_df[previous_df["Naam"] == name]

        new_tasks = len(set(current_tasks["Omschrijving middel"]) - set(prev_tasks["Omschrijving middel"]))
        old_tasks = prev_tasks[prev_tasks["Omschrijving middel"].isin(current_tasks["Omschrijving middel"])]

        old_oh_orders = ", ".join(old_tasks["OH-order"].astype(str).unique()) if not old_tasks.empty else ""

        comparison_data.append({
            "Naam": name,
            "Aantal nieuwe taken": new_tasks,
            "Aantal oude taken": len(old_tasks),
            "OH-orders (oude taken)": old_oh_orders
        })

    return pd.DataFrame(comparison_data)


def create_excel_file(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    writer.close()
    output.seek(0)
    return output


def create_custom_zip(
    everything_df=None,
    dict_per_name=None,
    aggregated_df=None,
    comparison_df=None,
    download_everything=False,
    download_per_name=False,
    download_aggregated=False,
    download_comparison=False,
    output_folder_name="output"
):
    output_zip = BytesIO()
    with zipfile.ZipFile(output_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        folder_path = f"{output_folder_name}/"

        if download_everything and everything_df is not None:
            file_bytes = create_excel_file(everything_df).getvalue()
            z.writestr(f"{folder_path}AllesBijElkaar.xlsx", file_bytes)

        if download_per_name and dict_per_name is not None:
            for name, df_group in dict_per_name.items():
                excel_bytes = create_excel_file(df_group).getvalue()
                valid_name = re.sub(r'[\\/*?:"<>|]', "_", str(name))
                z.writestr(f"{folder_path}{valid_name}.xlsx", excel_bytes)

        if download_aggregated and aggregated_df is not None:
            aggregated_bytes = create_excel_file(aggregated_df).getvalue()
            z.writestr(f"{folder_path}GegroepeerdOverzicht.xlsx", aggregated_bytes)

        if download_comparison and comparison_df is not None:
            comp_bytes = create_excel_file(comparison_df).getvalue()
            z.writestr(f"{folder_path}Vergelijking.xlsx", comp_bytes)

    output_zip.seek(0)
    return output_zip
