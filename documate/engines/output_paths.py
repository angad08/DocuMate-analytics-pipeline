"""
Small file helpers both engines use: output folder, output name, cleanup.
"""

import os
import time
from datetime import datetime

from documate.setup import messages


def make_output_folder(folder):
    """Create the output folder if it isn't there, saying so once."""
    if not os.path.exists(folder):
        print(messages.CREATING_FOLDER.format(folder=folder))
        os.makedirs(folder)
    return folder


def build_output_name(folder, prefix, timestamp_format):
    """
    Build the output file path, e.g.

        files\\output\\DocuMateX_BIRTH_REGISTRATION_13092026.docx

    The timestamp is worked out now, when this is called.

    The old scripts wrote the timestamp into a default argument, like:

        def generate(self, filename=f"...{datetime.now()}.docx")

    Python works out a default argument ONCE, when the file is first
    imported - not each time you call it. Z and O check for new records
    every 30 seconds, so every run reused the timestamp from when the
    program started and overwrote the previous document. This fixes that.
    """
    stamp = datetime.now().strftime(timestamp_format)
    filename = prefix + "_BIRTH_REGISTRATION_" + stamp + ".docx"
    return os.path.join(folder, filename)


def delete_temp_file(path, attempts=10, wait_seconds=1):
    """
    Delete a temp file, retrying while Word still has hold of it.

    Word lets go of its files a moment after Quit() returns, so deleting
    straight away fails. These temp files contain real personal data, so if
    we truly can't delete one we say so rather than quietly leaving it.
    """
    for _ in range(attempts):
        try:
            if os.path.exists(path):
                os.remove(path)
            return
        except OSError:
            time.sleep(wait_seconds)

    print(messages.CLEANUP_FAILED.format(path=path))
