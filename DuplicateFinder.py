"""Python module that searches for duplicate images. Prints all duplicates to terminal.
TODO delete unnecessary prints and dead code
TODO add multiprocessing to hashing function
TODO add time search functionality
BUG if there are fields with the same name, there dup will be reported th proper amount, but it will be the same fingerprint for each row!
NOTE In docs, remind it searches all photos in dir, so GROUP # will list dups for date range, bit not show all occurences of the dups.
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
    duplicate_dict, file_hash_dict = find_duplicates(photos)

    if (0 == len(duplicate_dict)):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + "[+] No duplicates detected!")
        exit_program(SUCCESS)

    ###########################################
    # Write the data to an Excel spreadhseet. #
    ###########################################
    with pd.ExcelWriter(abs_path_spreadsheet) as my_writer:
        df = dup_to_excel(abs_path_chat_history, duplicate_dict, file_hash_dict, date_range, my_writer)
        snitch(df, my_writer)

    exit_program(SUCCESS)



def snitch(group_df, my_writer):
    """Writes data to the Excel spreadsheet for individual uploaders of duplicate photos.
    
    Args:
        group_df (DataFrame): Contains all relevant info on unique duplicate groups to include: names, day, time, and photo paths.
        my_writer (ExcelWriter): Needed for .to_excel() to know where to write. Note: the writer should already be open.
    """
    #print(group_df)

    uploaders = group_df[UPLOADER].unique()

    #Filter out non-names from the column of names.
    uploaders = [i for i in uploaders if GROUP not in i]
    uploaders = [i for i in uploaders if "" != i]

    for name in uploaders:
        rows = group_df[group_df[UPLOADER] == name]
        # rows = rows.sort_values(["HASH"])
        # print(rows)
        total_dups_uploaded = len(rows)

        
        df = pd.DataFrame(columns=[DATE, TIME, FILE_PATH, HASH, ACCOMPLICES])
        totals_row = pd.DataFrame({DATE: [f"{name} uploaded {total_dups_uploaded} duplicates."], TIME: [""], FILE_PATH: [""], HASH: [""], ACCOMPLICES: [""]})
        blank_row = pd.DataFrame({DATE: [""], TIME: [""], FILE_PATH: [""], HASH: [""], ACCOMPLICES: [""]})
        df = pd.concat([df, rows])

        df = df.sort_values(by=[HASH])

        df = pd.concat([df, blank_row, totals_row])
        
        df.to_excel(excel_writer=my_writer, sheet_name=f"{name}", columns=[DATE, TIME, FILE_PATH, HASH, ACCOMPLICES], index=False)

    # print(group_df[UPLOADER].value_counts())



def dup_to_excel(chat_history, duplicate_dict, file_hash_dict, date_range, my_writer):
    """Clumps together a unique duplicate group into a data frame.
    Concatenates each group (data frame) together, and writes one large data frame to an Excel spreadsheet.

    Args:
        chat_history (str): Path to text file containig the WhatsApp chat history.
        duplicate_dict (dict): Contains all discovered duplicates w/ k/v being <duplicate hash>:<list of duplicate photo file paths>.
        file_hash_dict (dict): To identify the hash for each unique absolute path for dup photos, k/v being <abs_path>:hash.
        date_range (tuple): (start, end) Date range of photos we are interested in searching for.
        my_writer (ExcelWriter): Needed for .to_excel() to know where to write. Note: the writer should already be open.

    Returns:
        DataFrame: Contains all relevant info on duplicates including names, day, time, and photo paths.
    """

    #########################################################
    # Read in the chat history text and seperate by newline #
    #########################################################
    try:
        with open(chat_history, "r") as file:
            file_data = file.read()
    except OSError as oerr:
        print(f"[!] open({chat_history}): {oerr}")
        exit_program(ERROR)

    file_lines = file_data.splitlines()

    ###########################################
    # Isolate a single duplicat to work with. #
    ###########################################
    dup_group_total = 0 #Total unique duplicate groups.
    dup_total = 0 #Total of duplicate photos that exist.

    df = pd.DataFrame(columns=[UPLOADER, DATE, TIME, FILE_PATH, HASH])
    dups_list = duplicate_dict.values() #dups_list is a list of lists.

    dup_total_prev = dup_total

    for dup in dups_list: #dup is a list.
        #print("Duplicate:")
        # In this loop we are dealing with a single duplicate.
        # if dup_total_prev 
        dup_group_total += 1
        title_row = pd.DataFrame({UPLOADER: [f"{GROUP}{dup_group_total}"], DATE: [""], TIME: [""], FILE_PATH: [""], HASH: [""]})
        df = pd.concat([df, title_row])

        for photo_abs_path in dup: #photo_path is a string.
            # print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + f"\t{photo_abs_path}")
            _, photo_file_name = os.path.split(photo_abs_path)

            #Find line (there should only be 1, but this will handle multipe occurences) where the file name of the dup appears.
            for line in file_lines:
                line = line.replace("\u202f", " ").replace("\u200e", "") #Get rid of Unicode narrow space and LTRM.
                #Found the line where the photo was uploaded.
                if photo_file_name in line:
                    #print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + f"\t\t{line}")

                    #######################################################
                    # Pick apart the chat history line for relevant info. #
                    #######################################################
                    pattern = re.compile(r"\[(\d{1,2}/\d{1,2}/\d{2}),\s*([^\]]+)\]\s*~?\s*(.*?):.*?<attached:\s*(.*?)>")

                    match = pattern.search(line)

                    if match:
                        upload_date = match.group(1)
                        upload_time = match.group(2)
                        uploader    = match.group(3)
                        photo_path  = match.group(4)

                        ###################################################################################################################
                        # Determine if the photo is in our date range, if we have one.                                                    #
                        # Note that ouu dict args have all photos, as we want to compare our photos in our date range against ALL photos. #
                        # I.e., the date range isn't to limit our input search, but to limit our output to the Excel spreadsheet.         #
                        ###################################################################################################################
                        if not date_range:
                            pass #No date range
                        else:
                            start_date, end_date = date_range

                            try:
                                upload_date_formatted = pd.to_datetime(upload_date)
                            except (pd.errors.ParserError, ValueError) as err:
                                print(colorama.Fore.RED + colorama.Style.BRIGHT + f"[X] Parsing upload date failed! {err}" + colorama.Style.RESET_ALL)
                                exit_program(ERROR)
                            else:
                                #Outside the range, do not add to dataframe that will be written to the Excel spreadsheet.
                                if not (start_date <= upload_date_formatted <= end_date):
                                    continue


                        #######################################
                        # Store into a data frame and append. #
                        #######################################
                        hash_val = file_hash_dict.get(photo_abs_path)
                        new_row = pd.DataFrame([{UPLOADER: uploader, DATE: upload_date, TIME: upload_time, FILE_PATH: photo_path, HASH: hash_val}])
                        df      = pd.concat([df, new_row])
                        dup_total += 1
                    else:
                        print(colorama.Fore.YELLOW + colorama.Style.BRIGHT + "[!] Regex search did not work!")

            #print(f"{df}")
        #print() #TODO delete this line
        blank_rows = pd.DataFrame({UPLOADER: ["", ""], DATE: ["", ""], TIME: ["", ""], FILE_PATH: ["", ""], HASH: ["", ""]})
        df = pd.concat([df, blank_rows])

        ############################################
        # Write all duplicates to the spreadsheet. #
        ############################################
    #print(f"{df}")
    df.to_excel(excel_writer=my_writer, sheet_name="Duplicate Photos", columns=[UPLOADER, DATE, TIME, FILE_PATH, HASH], index=False)

    print(colorama.Fore.GREEN + colorama.Style.NORMAL + f"[+] Total unique duplicate groups:   {dup_group_total}\n[+] Total duplicate photos detected: {dup_total}" + colorama.Style.RESET_ALL)

    return df



def find_duplicates(photos):
    """Finds duplicate photos in a list.

    Args:
        photos (list): All photos that need to be analyzed for existence of duplicates.

    Returns:
        tuple: First dict contains all discovered duplicates w/ k/v being <duplicate hash>:<list of duplicate photo file paths>.
               Second dict contains abs_path:hash for each unique absolute path, used in dup_to_excel().
    """

    hash_dict = {} #local dict hash:photo
    duplicate_dict = {} #ret dict hash:<list of photos>
    file_hash_dict = {} #ret dict file_abs_path:hash
    duplicate_dict = defaultdict(list)

    for photo in photos:
        with open(photo, "rb") as f:
            digest = hashlib.file_digest(f, "md5") #TODO use faster hash like blake3 ... using this for convenient file_digest methood.
            hash = digest.hexdigest()

            #For use in dup_to_excel(), update this dictionary.
            file_hash_dict[photo] = hash
            
            #A duplicate photo is detected.
            if hash in hash_dict:
                #If this is the first time the duplicate is hit, update duplicate dict w/ original file too.
                if 0 == len(duplicate_dict[hash]):
                    duplicate_dict[hash].append(hash_dict[hash])

                #Update duplicate dict w/ the duplicate.
                duplicate_dict[hash].append(photo)

            #Photo is not a duplicate; update hash dictionary with new hash.
            else:
                hash_dict[hash] = photo

    #Print out existing duplicates to terminal. #TODO delete??
    for dup_hash in duplicate_dict.keys():
        #print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + f"[*] Duplicate file detected {len(duplicate_dict[dup_hash])} times:")
        for file in duplicate_dict[dup_hash]:
            #print(f"\t{file}")
            pass #TODO delete?

    #print() #TODO delete this line
    return (duplicate_dict, file_hash_dict)



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
        spreadsheet_name = os.path.join(root, ".xlsx")

    abs_path_spreadsheet = os.path.abspath(spreadsheet_name)

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
