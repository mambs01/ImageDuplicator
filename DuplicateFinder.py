"""Python module that searches for duplicate images. Prints all duplicates to terminal.
TODO delete unnecessary prints and dead code
TODO add multiprocessing to hashing function
TODO add time search functionality
BUG if there are fields with the same name, there dup will be reported th proper amount, but it will be the same fingerprint for each row!
NOTE In docs, remind it searches all photos in dir, so GROUP # will list dups for date range, bit not show all occurences of the dups.
BUG total unique duplicate groups # not accurate for date range, excel prints blank groups.
I.e., Excel won't show dups outside of the range, or if running for older date range, will search even new photos, so data might not add up to old sheet (more dups detected in newer run)

Next steps: verify duplicates are accurate, and the false positive is no longer there.
What should the Excel sheet look like... how much info is too much?
"""

import os
import re
import argparse
import colorama
import filetype
import hashlib

#from PIL import Image
from collections import defaultdict
import pandas as pd
#from collections import Counter

#Data Frame Columns
UPLOADER    = "UPLOADER"
DATE        = "DATE"
TIME        = "TIME"
FILE_PATH   = "FILE PATH"
HASH        = "HASH"
ACCOMPLICES = "ACCOMPLICES"

#Program exit codes
SUCCESS = 0
ERROR   = 1

#MISC
GROUP = "GROUP #"



def main():
    """Core logic of the script, and where the script starts."""

    ################################################################
    # Get command line arguments if they exist, and verify input. #
    ################################################################
    parser = argparse.ArgumentParser(description="Searches for duplicate images. Prints all duplicates to terminal.", formatter_class=argparse.ArgumentDefaultsHelpFormatter) #TODO Update desc.
    parser.add_argument('-p', '--photos', default=os.getcwd(), type=str, metavar="<path to folder w/ photos>",help="Path to the folder where the photos to be searched are located.")
    parser.add_argument('-c', '--chathistory', default="chat_history_full.txt", type=str, metavar="<path to chat history file>", help="Path to the WhatsApp chat history text file.")
    parser.add_argument('-o', '--output', default="duplicates.xlsx", type=str, metavar="<desired Excel spreadsheet name>", help="Choose the name of the Excel spreadsheet file that is generated on output. Can also state the absolute path of the file.")
    parser.add_argument('-s', '--start_day', type=str, metavar="<start day>", help="Optionally limit search of photos from start day to end day, -s and -e must exist together.")
    parser.add_argument('-e', '--end_day', type=str, metavar="<start day>", help="Optionally limit search of photos from start day to end day, -s and -e must exist together.")

    args             = parser.parse_args()
    photo_folder     = args.photos
    chat_history     = args.chathistory
    spreadsheet_name = args.output

    abs_path_photo_folder, abs_path_chat_history = validate_paths(photo_folder, chat_history)
    abs_path_spreadsheet = validate_ext(spreadsheet_name)

    date_range = validate_dates(args.start_day, args.end_day)
    
    ####################################################################
    # Filter out non image files; don't waste time on non-image files. #
    ####################################################################
    photos = extract_photos(abs_path_photo_folder)
    #print(len(photos))

    ################################
    # Search for duplicate images. #
    ################################
    duplicate_dict = find_duplicates(photos)

    if (0 == len(duplicate_dict)):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + "[+] No duplicates detected!")
        exit_program(SUCCESS)

    ###########################################
    # Write the data to an Excel spreadhseet. #
    ###########################################
    with pd.ExcelWriter(abs_path_spreadsheet) as my_writer:
        df = dup_to_excel(abs_path_chat_history, duplicate_dict, date_range, my_writer)
        snitch(df, my_writer)

    exit_program(SUCCESS)



def snitch(group_df, my_writer):
    """Writes a per-person sheet to Excel for each uploader who submitted duplicate photos.

    Args:
        group_df (DataFrame): The full DataFrame returned by dup_to_excel(), containing both
                              structural rows (GROUP headers, blank separators) and data rows.
        my_writer (ExcelWriter): Needed for .to_excel() to know where to write. Note: the writer should already be open.
    """

    #Filter uploader names from the UPLOADER column, excluding GROUP headers and blank separators.
    uploaders = group_df[UPLOADER].unique()
    uploaders = [u for u in uploaders if (pd.notna(u) and ("" != u) and (GROUP not in str(u)))]

    for name in uploaders:
        rows = group_df[group_df[UPLOADER] == name].sort_values(by=[HASH])
        total_dups_uploaded = len(rows)

        blank_row  = pd.DataFrame([{DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""}])
        totals_row = pd.DataFrame([{DATE: f"{name} uploaded {total_dups_uploaded} duplicates.", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""}])
        df = pd.concat([rows, blank_row, totals_row], ignore_index=True)

        df.to_excel(excel_writer=my_writer, sheet_name=name, columns=[DATE, TIME, FILE_PATH, HASH, ACCOMPLICES], index=False)



def dup_to_excel(chat_history, duplicate_dict, date_range, my_writer):
    """Builds duplicate-group data and writes the summary sheet to Excel.

    For each duplicate group detected on disk, looks up matching chat history entries
    by filename, applies any date filter, and collects the results. Groups with no
    chat matches (or no matches within the date range) are skipped entirely so the
    spreadsheet never has empty GROUP rows. Accomplices are filled in per-row before
    writing.

    Args:
        chat_history (str): Absolute path to the WhatsApp chat history text file.
        duplicate_dict (dict): k/v being <hash>:<list of abs_paths> from find_duplicates().
        date_range (tuple | None): (start, end) pandas Timestamps, or None for no filter.
        my_writer (ExcelWriter): Open ExcelWriter to write the summary sheet into.

    Returns:
        DataFrame: All data rows (plus structural header/blank rows) written to the sheet.
                   Passed to snitch() to build per-person sheets.
    """

    ############################################################
    # Parse chat history once into a filename-keyed lookup.    #
    # This replaces scanning all lines for every single photo. #
    ############################################################
    chat_lookup = build_chat_lookup(chat_history)

    #####################################################################
    # For each duplicate group, collect matching chat rows, date-filter,#
    # and skip the group entirely if nothing survives the filter.        #
    #####################################################################
    all_groups = []  #list of group_rows lists; only non-empty groups after filtering.

    for hash_val, abs_paths in duplicate_dict.items():
        group_rows = []

        for abs_path in abs_paths:
            filename = os.path.basename(abs_path)
            entries  = chat_lookup.get(filename, [])

            if (not entries):
                print(colorama.Fore.YELLOW + colorama.Style.BRIGHT + f"[!] {filename} not found in chat history, skipping." + colorama.Style.RESET_ALL)
                continue

            for entry in entries:
                ###################################################################################################################
                # Determine if the photo is in our date range, if we have one.                                                    #
                # Note that our dict args have all photos, as we want to compare our photos in our date range against ALL photos. #
                # I.e., the date range isn't to limit our input search, but to limit our output to the Excel spreadsheet.        #
                ###################################################################################################################
                if (date_range):
                    start_date, end_date = date_range

                    try:
                        upload_date_formatted = pd.to_datetime(entry["date"])
                    except (pd.errors.ParserError, ValueError) as err:
                        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Parsing upload date failed! {err}" + colorama.Style.RESET_ALL)
                        exit_program(ERROR)

                    if (not (start_date <= upload_date_formatted <= end_date)):
                        continue  #Outside the range, skip this photo.

                group_rows.append({
                    UPLOADER:    entry["uploader"],
                    DATE:        entry["date"],
                    TIME:        entry["time"],
                    FILE_PATH:   entry["file_path"],
                    HASH:        hash_val,
                    ACCOMPLICES: "",  #Filled in below once we know all uploaders in this group.
                })

        if (not group_rows):
            continue  #No chat matches or all filtered by date; skip this group entirely.

        #######################################################################################
        # Fill ACCOMPLICES: unique other uploaders in this group, excluding the row's own  #
        # uploader. Name-based deduplication means two distinct people with the same name  #
        # will incorrectly collapse into one (noted in README under "Notes").              #
        #######################################################################################
        unique_uploaders = list(dict.fromkeys(r[UPLOADER] for r in group_rows))
        for row in group_rows:
            others = [u for u in unique_uploaders if (u != row[UPLOADER])]
            row[ACCOMPLICES] = ", ".join(others)

        all_groups.append(group_rows)

    ##########################################################################
    # Build the flat row list that becomes the "Duplicate Photos" sheet.     #
    # Structural rows (GROUP headers and blank separators) are added here,   #
    # only for groups that have actual data.                                  #
    ##########################################################################
    flat_rows = []
    for i, group_rows in enumerate(all_groups, 1):
        flat_rows.append({UPLOADER: f"{GROUP}{i}", DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""})
        flat_rows.extend(group_rows)
        flat_rows.append({UPLOADER: "", DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""})
        flat_rows.append({UPLOADER: "", DATE: "", TIME: "", FILE_PATH: "", HASH: "", ACCOMPLICES: ""})

    dup_group_total = len(all_groups)
    dup_total       = sum(len(g) for g in all_groups)

    df = pd.DataFrame(flat_rows, columns=[UPLOADER, DATE, TIME, FILE_PATH, HASH, ACCOMPLICES])
    df.to_excel(excel_writer=my_writer, sheet_name="Duplicate Photos", columns=[UPLOADER, DATE, TIME, FILE_PATH, HASH], index=False)

    print(colorama.Fore.GREEN + colorama.Style.NORMAL + f"[+] Total unique duplicate groups:   {dup_group_total}\n[+] Total duplicate photos detected: {dup_total}" + colorama.Style.RESET_ALL)

    return df



def find_duplicates(photos):
    """Finds duplicate photos in a list.

    Args:
        photos (list): All photos that need to be analyzed for existence of duplicates.

    Returns:
        dict: k/v being <duplicate hash>:<list of duplicate photo file paths>.
    """

    hash_dict = {}  #local dict hash:first_seen_abs_path
    duplicate_dict = defaultdict(list)  #ret dict hash:[abs_path, ...]

    for photo in photos:
        with open(photo, "rb") as f:
            digest = hashlib.file_digest(f, "md5") #TODO use faster hash like blake3 ... using this for convenient file_digest method.
            hash = digest.hexdigest()

            #A duplicate photo is detected.
            if (hash in hash_dict):
                #If this is the first time the duplicate is hit, add the original file too.
                if (0 == len(duplicate_dict[hash])):
                    duplicate_dict[hash].append(hash_dict[hash])

                #Add the duplicate.
                duplicate_dict[hash].append(photo)

            #Photo is not a duplicate; store it in case a duplicate appears later.
            else:
                hash_dict[hash] = photo

    return duplicate_dict



def build_chat_lookup(chat_history_path):
    """Parses the WhatsApp chat history file once into a filename-keyed lookup dict.

    Args:
        chat_history_path (str): Absolute path to the WhatsApp chat history text file.

    Returns:
        dict: {whatsapp_filename: [{"uploader": str, "date": str, "time": str, "file_path": str}, ...]}
              A single filename can have multiple entries if it appears on more than one chat line.
    """
    try:
        with open(chat_history_path, "r") as f:
            file_lines = f.read().splitlines()
    except OSError as oerr:
        print(f"[!] open({chat_history_path}): {oerr}")
        exit_program(ERROR)

    pattern = re.compile(r"\[(\d{1,2}/\d{1,2}/\d{2}),\s*([^\]]+)\]\s*~?\s*(.*?):.*?<attached:\s*(.*?)>")
    lookup = defaultdict(list)

    for line in file_lines:
        line = line.replace(" ", " ").replace("‎", "")  #Remove Unicode narrow space and LTRM.
        match = pattern.search(line)
        if match:
            lookup[match.group(4)].append({
                "uploader":  match.group(3),
                "date":      match.group(1),
                "time":      match.group(2),
                "file_path": match.group(4),
            })

    return lookup



def extract_photos(photo_folder):
    """Recursively searches photo_folder and filters out non-images.
    For a list of all supported image file types see:
        1) https://pypi.org/project/filetype/
        2) https://github.com/h2non/filetype.py/blob/master/filetype/types/image.py

    Args:
        photo_folder (str): Folder path that will be recursively searched for all photos.

    Returns:
        list: Absolute paths of all photos in the searched folder(s).
    """
    photos = []

    for root, _, files in os.walk(photo_folder):
        for file in files:
            file_path = os.path.join(root, file)
            file_abs_path = os.path.abspath(file_path)

            try:
                is_image = filetype.is_image(file_abs_path)
            except (TypeError, IsADirectoryError) as err:
                print(colorama.Fore.YELLOW + colorama.Style.BRIGHT + f"[!] Filetype for {file_abs_path} not supported! {err}" + colorama.Style.RESET_ALL)
            else:
                if True == is_image:
                    photos.append(file_abs_path)

    return photos



def validate_dates(start_day, end_day):
    """Determines if the user supplied a date range to search photo search.
    Additionally, validates the input to make sure it is in proper format and a valid date.

    Args:
        start_day (str): Earliest date to search for in range.
        end_day (str): Latest date to search for in range.

    Returns:
        tuple | None: tuple of pandas.to_datetime (start, end) which can be various return types, or None if both args are not supplied.
    """

    ####################################
    # Check if both args are supplied. #
    ####################################

    #Both args not supplied, user does not care about date range; user is searching all photos.
    if ((not start_day) and (not end_day)):
        print(colorama.Fore.GREEN + colorama.Style.NORMAL + "[+] No date range detected, searching all photos." + colorama.Style.RESET_ALL)
        return None
    #Both args supplied, user wants to search in a date range.
    elif (start_day and end_day):
        try:
            start_day_formatted = pd.to_datetime(start_day)
            end_day_formatted   = pd.to_datetime(end_day)
        except (pd.errors.ParserError, ValueError) as err:
            print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Parsing start/end dates failed. Enter in MM/DD/YY format. {err}" + colorama.Style.RESET_ALL)
            exit_program(ERROR)
        else:
            print(colorama.Fore.GREEN + colorama.Style.NORMAL + f"[+] Searching photos from {start_day_formatted} to {end_day_formatted}" + colorama.Style.RESET_ALL)
            return(start_day_formatted, end_day_formatted)
    else:
        print(colorama.Fore.RED + colorama.Style.BRIGHT + "[X] Forgot to add start or end date! Must use both the -s and -e options! Use -h for help!" + colorama.Style.RESET_ALL)
        exit_program(ERROR)


    
    


def validate_paths(photo_folder, chat_history):
    """ Convert the file and directory args to absolute paths and verify they exist.

    Args:
        photo_folder (str): Path to the folder where the photos to be searched are at.
        chat_history (str): Path to the WhatsApp chat history text file.

    Returns:
        tuple: Absolute paths for (photo_folder, chat_history).

    
    """
    photo_folder_abs_path = os.path.abspath(photo_folder)
    chat_history_abs_path = os.path.abspath(chat_history)

    #Photo folder is invalid.
    if not (os.path.isdir(photo_folder_abs_path)):
        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X]: {photo_folder_abs_path} does not exist!")
        exit_program(ERROR)
    #Chat history file is invalid.
    elif not (os.path.isfile(chat_history_abs_path)):
        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X]: {chat_history_abs_path} does not exist!")
        exit_program(ERROR)
    #Both files are valid.
    else:
        print(colorama.Fore.GREEN + f"[+] Using photos located in:            {photo_folder_abs_path}")
        print(f"[+] Using chat history file located at: {chat_history_abs_path}"  + colorama.Style.RESET_ALL)

    return (photo_folder_abs_path, chat_history_abs_path)



def validate_ext(spreadsheet_name):
    """Appends the .xlsx file extension onto the name of the outputted spreadsheet.
    The file extensio nis required for Pandas Excel functions.

    Args:
        spreadsheet_name (str): Name of the script's outputted spreadhseet.

    Returns:
        str: Absolute path of spreadsheet_name.
    """

    root, ext = os.path.splitext(spreadsheet_name)

    if (".xlsx" == ext):
        pass
    else:
        # Strip any trailing path separators before appending the extension,
        # so that e.g. "output/" becomes "output.xlsx" not "output/.xlsx".
        spreadsheet_name = root.rstrip(os.sep) + ".xlsx"

    abs_path_spreadsheet = os.path.abspath(spreadsheet_name)

    if (os.path.isdir(abs_path_spreadsheet)):
        print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Output path is a directory, not a file: {abs_path_spreadsheet}" + colorama.Style.RESET_ALL)
        exit_program(ERROR)

    print(colorama.Fore.GREEN + colorama.Style.NORMAL + f"[+] Writing spreadsheet to:             {abs_path_spreadsheet}" + colorama.Style.RESET_ALL)

    return abs_path_spreadsheet



def exit_program(exit_val):
    """Exits the program and resets colorama.
    
    Args:
        exit_val (int): The exit value of the program.
    """

    print(colorama.Fore.GREEN + "[+] Exiting script, goodbye!" + colorama.Style.RESET_ALL)
    raise SystemExit(exit_val)



if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        exit_program(SUCCESS)
