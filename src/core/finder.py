import hashlib
import os
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import colorama
import filetype


def _hash_file(path):
    """Read and hash a single file. Called in a thread pool by find_duplicates()."""
    with open(path, "rb") as f:
        return (path, hashlib.file_digest(f, "blake2b").hexdigest())


def find_duplicates(photos):
    """Hash every photo and return only those that appear more than once.

    Args:
        photos (list): Absolute paths to all photos to analyze.

    Returns:
        dict: {hash: [abs_path, ...]} containing only duplicate groups.
    """
    hash_dict      = {}
    duplicate_dict = defaultdict(list)

    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(_hash_file, p) for p in photos]
        results = [f.result() for f in as_completed(futures)]

    for path, hash_val in results:
        if hash_val in hash_dict:
            if len(duplicate_dict[hash_val]) == 0:
                duplicate_dict[hash_val].append(hash_dict[hash_val])
            duplicate_dict[hash_val].append(path)
        else:
            hash_dict[hash_val] = path

    return duplicate_dict


def extract_photos(photo_folder):
    """Recursively search photo_folder and return absolute paths of all image files.

    Args:
        photo_folder (str): Root folder to search.

    Returns:
        list: Absolute paths of every image found.
    """
    photos = []

    for root, _, files in os.walk(photo_folder):
        for file in files:
            file_abs_path = os.path.abspath(os.path.join(root, file))
            try:
                is_image = filetype.is_image(file_abs_path)
            except (TypeError, IsADirectoryError) as err:
                print(colorama.Fore.YELLOW + colorama.Style.BRIGHT +
                      f"[!] Filetype for {file_abs_path} not supported! {err}" +
                      colorama.Style.RESET_ALL)
            else:
                if is_image:
                    photos.append(file_abs_path)

    return photos
