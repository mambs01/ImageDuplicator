import os

import colorama
import pandas as pd

from core.chat_parser import build_chat_lookup
from core.constants import (ACCOMPLICES, DATE, FILE_PATH, GROUP, HASH, TIME,
                             UPLOADER)


def snitch(group_df, my_writer):
    """Write a per-person sheet for each uploader who submitted duplicate photos.

    Args:
        group_df (DataFrame): Full DataFrame from dup_to_excel(), including structural rows.
        my_writer (ExcelWriter): Open ExcelWriter to write into.
    """
    uploaders = group_df[UPLOADER].unique()
    uploaders = [u for u in uploaders if pd.notna(u) and u != "" and GROUP not in str(u)]

    for name in uploaders:
        rows      = group_df[group_df[UPLOADER] == name].sort_values(by=[HASH])
        total     = len(rows)
        blank_row = pd.DataFrame([{DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""}])
        total_row = pd.DataFrame([{DATE: f"{name} uploaded {total} duplicates.",
                                   TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""}])
        df = pd.concat([rows, blank_row, total_row], ignore_index=True)
        df.to_excel(excel_writer=my_writer, sheet_name=name,
                    columns=[DATE, TIME, FILE_PATH, HASH, ACCOMPLICES], index=False)


def dup_to_excel(chat_history, duplicate_dict, date_range, my_writer):
    """Build duplicate-group data and write the summary sheet to Excel.

    For each duplicate group, looks up matching chat history entries by filename,
    applies any date filter, and skips groups with no matches. Passes the resulting
    DataFrame to snitch() to build per-person sheets.

    Args:
        chat_history (str): Absolute path to the WhatsApp chat history .txt file.
        duplicate_dict (dict): {hash: [abs_paths]} from find_duplicates().
        date_range (tuple | None): (start, end) pandas Timestamps, or None for no filter.
        my_writer (ExcelWriter): Open ExcelWriter to write the summary sheet into.

    Returns:
        DataFrame: All rows written to the sheet (including structural rows).
    """
    chat_lookup = build_chat_lookup(chat_history)
    all_groups  = []

    for hash_val, abs_paths in duplicate_dict.items():
        group_rows = []

        for abs_path in abs_paths:
            filename = os.path.basename(abs_path)
            entries  = chat_lookup.get(filename, [])

            if not entries:
                print(colorama.Fore.YELLOW + colorama.Style.BRIGHT +
                      f"[!] {filename} not found in chat history, skipping." +
                      colorama.Style.RESET_ALL)
                continue

            for entry in entries:
                if date_range:
                    start_date, end_date = date_range
                    try:
                        upload_date = pd.to_datetime(entry["date"])
                    except (pd.errors.ParserError, ValueError) as err:
                        print(colorama.Fore.RED + colorama.Style.BRIGHT +
                              f"[X] Parsing upload date failed! {err}" +
                              colorama.Style.RESET_ALL)
                        raise SystemExit(1)

                    if not (start_date <= upload_date <= end_date):
                        continue

                group_rows.append({
                    UPLOADER:    entry["uploader"],
                    DATE:        entry["date"],
                    TIME:        entry["time"],
                    FILE_PATH:   entry["file_path"],
                    HASH:        hash_val,
                    ACCOMPLICES: "",
                })

        if not group_rows:
            continue

        unique_uploaders = list(dict.fromkeys(r[UPLOADER] for r in group_rows))
        for row in group_rows:
            others = [u for u in unique_uploaders if u != row[UPLOADER]]
            row[ACCOMPLICES] = ", ".join(others)

        all_groups.append(group_rows)

    flat_rows = []
    for i, group_rows in enumerate(all_groups, 1):
        flat_rows.append({UPLOADER: f"{GROUP}{i}", DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""})
        flat_rows.extend(group_rows)
        flat_rows.append({UPLOADER: "", DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""})
        flat_rows.append({UPLOADER: "", DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""})

    dup_group_total = len(all_groups)
    dup_total       = sum(len(g) for g in all_groups)

    df = pd.DataFrame(flat_rows, columns=[UPLOADER, DATE, TIME, FILE_PATH, HASH, ACCOMPLICES])
    df.to_excel(excel_writer=my_writer, sheet_name="Duplicate Photos",
                columns=[UPLOADER, DATE, TIME, FILE_PATH, HASH], index=False)

    print(colorama.Fore.GREEN + colorama.Style.NORMAL +
          f"[+] Total unique duplicate groups:   {dup_group_total}\n"
          f"[+] Total duplicate photos detected: {dup_total}" +
          colorama.Style.RESET_ALL)

    return df
