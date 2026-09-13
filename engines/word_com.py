"""
Word's magic numbers, given names, plus a safe way to import pywin32.

Word's COM interface takes numbers for everything. 0 means "send to a new
document", 7 means "section break, next page". Giving them names means
mailmerge.py reads as words instead of digits.
"""

# Where the merge result goes: a new document.
SEND_TO_NEW_DOCUMENT = 0

# When closing a document: don't save.
DO_NOT_SAVE_CHANGES = 0

# Save format: .docx
FORMAT_DOCX = 16

# Break type: section break, starts on the next page.
SECTION_BREAK_NEXT_PAGE = 7

# Mail merge document type: form letters.
FORM_LETTERS = 0

# Macro security: force-disable. This also stops Word asking you to confirm
# the data source, which would otherwise freeze an unattended run.
DISABLE_MACROS = 3


def load_word():
    """
    Load pywin32 and hand back what we need, or explain why we can't.

    The old scripts did "import win32com" at the top of the file and called
    sys.exit(1) if it failed. That meant the whole file couldn't even be
    opened on a non-Windows machine, and simply importing it could kill the
    program. We import it here instead, only when Mail Merge actually runs.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "Word Mail Merge needs pywin32 and Microsoft Word, on Windows.\n"
            "Install it with: pip install pywin32"
        )

    return pythoncom, win32com.client
