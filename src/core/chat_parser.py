import re
from collections import defaultdict

import colorama


def build_chat_lookup(chat_history_path):
    """Parse the WhatsApp chat history file into a filename-keyed lookup dict.

    Args:
        chat_history_path (str): Absolute path to the WhatsApp chat history .txt file.

    Returns:
        dict: {filename: [{"uploader", "date", "time", "file_path"}, ...]}
    """
    try:
        with open(chat_history_path, "r") as f:
            file_lines = f.read().splitlines()
    except OSError as oerr:
        print(colorama.Fore.RED + colorama.Style.BRIGHT +
              f"[X] Could not open chat history: {oerr}" +
              colorama.Style.RESET_ALL)
        raise SystemExit(1)

    pattern = re.compile(
        r"\[(\d{1,2}/\d{1,2}/\d{2}),\s*([^\]]+)\]\s*~?\s*(.*?):.*?<attached:\s*(.*?)>"
    )
    lookup = defaultdict(list)

    for line in file_lines:
        line = line.replace(" ", " ").replace("‎", "")
        match = pattern.search(line)
        if match:
            lookup[match.group(4)].append({
                "uploader":  match.group(3),
                "date":      match.group(1),
                "time":      match.group(2),
                "file_path": match.group(4),
            })

    return lookup
