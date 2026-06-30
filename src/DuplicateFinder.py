"""
DuplicateFinder — CLI entry point.

Usage:
    python src/DuplicateFinder.py -p <photos> -c <chat> -o <output> [-s <start> -e <end>]
"""

import argparse
import os
import re

import colorama
import pandas as pd

from core.constants import ERROR, SUCCESS
from core.finder import extract_photos, find_duplicates
from core.spreadsheet import dup_to_excel, snitch


def main():
    ################################################################
    # Get command line arguments and verify input.                 #
    ################################################################
    parser = argparse.ArgumentParser(description="Searches for duplicate images and writes an Excel report.", formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('-p', '--photos',     default=os.getcwd(),              type=str, metavar="<path>", help="Folder containing photos to search.")
    parser.add_argument('-c', '--chathistory',default="chat_history_full.txt",  type=str, metavar="<path>", help="WhatsApp chat history .txt file.")
    parser.add_argument('-o', '--output',     default="duplicates.xlsx",        type=str, metavar="<path>", help="Output Excel spreadsheet path/name.")
    parser.add_argument('-s', '--start_day',  type=str, metavar="<MM/DD/YY>",  help="Start of date range (must pair with -e).")
    parser.add_argument('-e', '--end_day',    type=str, metavar="<MM/DD/YY>",  help="End of date range (must pair with -s).")

    args             = parser.parse_args()
    photo_folder     = args.photos
    chat_history     = args.chathistory
    spreadsheet_name = args.output

    abs_path_photo_folder, abs_path_chat_history = validate_paths(photo_folder, chat_history)
    abs_path_spreadsheet = validate_ext(spreadsheet_name)
    date_range           = validate_dates(args.start_day, args.end_day)

    ####################################################################
    # Filter out non-image files.                                      #
    ####################################################################
    photos = extract_photos(abs_path_photo_folder)

    ###########################################################################
    # Pre-flight checks.                                                      #
    #                                                                         #
    # 1. OS copy suffixes like " (1)" in a filename mean the archive was      #
    #    downloaded more than once. The chat history never records these       #
    #    suffixes, so affected files cannot be matched. Download a fresh copy  #
    #    and rerun.                                                            #
    #                                                                         #
    # 2. Duplicate base filenames across subdirectories mean the chat history  #
    #    lookup cannot tell the two files apart. Merge into one flat folder    #
    #    and rerun.                                                            #
    ###########################################################################
    copy_suffix    = re.compile(r"\(\d+\)\.[^.]+$")  # matches OS copy suffixes at end of filename, e.g. (1).jpg or (2).png
    seen_basenames = {}
    for path in photos:
        base = os.path.basename(path)
        if copy_suffix.search(base):
            print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] OS copy suffix detected: {path}\n    Download a fresh copy and rerun." + colorama.Style.RESET_ALL)
            exit_program(ERROR)

        if base in seen_basenames:
            print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Duplicate filename across subdirectories:\n      {seen_basenames[base]}\n      {path}\n    Merge all photos into one flat folder and rerun." + colorama.Style.RESET_ALL)
            exit_program(ERROR)

        seen_basenames[base] = path

    ################################
    # Search for duplicate images. #
    ################################
    duplicate_dict = find_duplicates(photos)

    if len(duplicate_dict) == 0:
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + "[+] No duplicates detected!")
        exit_program(SUCCESS)

    ###########################################
    # Write the data to an Excel spreadsheet. #
    ###########################################
    with pd.ExcelWriter(abs_path_spreadsheet) as my_writer:
        df = dup_to_excel(abs_path_chat_history, duplicate_dict, date_range, my_writer)
        snitch(df, my_writer)

    exit_program(SUCCESS)


def validate_dates(start_day, end_day):
    """Validate the optional date range arguments.

    Returns:
        tuple | None: (start, end) pandas Timestamps, or None if no range given.
    """
    if not start_day and not end_day:
        print(colorama.Fore.GREEN + colorama.Style.NORMAL + "[+] No date range detected, searching all photos." + colorama.Style.RESET_ALL)
        return None

    if start_day and end_day:
        try:
            start = pd.to_datetime(start_day)
            end   = pd.to_datetime(end_day)
        except (pd.errors.ParserError, ValueError) as err:
            print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Parsing start/end dates failed. Enter in MM/DD/YY format. {err}" + colorama.Style.RESET_ALL)
            exit_program(ERROR)
        else:
            print(colorama.Fore.GREEN + colorama.Style.NORMAL + f"[+] Searching photos from {start} to {end}" + colorama.Style.RESET_ALL)
            return (start, end)

    print(colorama.Fore.RED + colorama.Style.BRIGHT + "[X] Must use both -s and -e together. Use -h for help." + colorama.Style.RESET_ALL)
    exit_program(ERROR)


def validate_paths(photo_folder, chat_history):
    """Convert paths to absolute and verify they exist.

    Returns:
        tuple: (abs_photo_folder, abs_chat_history)
    """
    photo_folder_abs = os.path.abspath(photo_folder)
    chat_history_abs = os.path.abspath(chat_history)

    if not os.path.isdir(photo_folder_abs):
        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] {photo_folder_abs} does not exist!")
        exit_program(ERROR)
    elif not os.path.isfile(chat_history_abs):
        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] {chat_history_abs} does not exist!")
        exit_program(ERROR)
    else:
        print(colorama.Fore.GREEN + f"[+] Using photos located in:            {photo_folder_abs}\n[+] Using chat history file located at: {chat_history_abs}" + colorama.Style.RESET_ALL)

    return (photo_folder_abs, chat_history_abs)


def validate_ext(spreadsheet_name):
    """Ensure the spreadsheet name ends with .xlsx.

    Returns:
        str: Absolute path of the output spreadsheet.
    """
    root, ext = os.path.splitext(spreadsheet_name)

    if ext != ".xlsx":
        spreadsheet_name = root.rstrip(os.sep) + ".xlsx"

    abs_path = os.path.abspath(spreadsheet_name)

    if os.path.isdir(abs_path):
        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Output path is a directory, not a file: {abs_path}" + colorama.Style.RESET_ALL)
        exit_program(ERROR)

    print(colorama.Fore.GREEN + colorama.Style.NORMAL + f"[+] Writing spreadsheet to:             {abs_path}" + colorama.Style.RESET_ALL)

    return abs_path


def exit_program(exit_val):
    """Print goodbye message and exit."""
    print(colorama.Fore.GREEN + "[+] Exiting program, goodbye!" + colorama.Style.RESET_ALL)
    raise SystemExit(exit_val)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        exit_program(SUCCESS)
