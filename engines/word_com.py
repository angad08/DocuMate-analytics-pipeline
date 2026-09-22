"""
Word COM constants, and a helper to load pywin32.

Word's COM interface uses numbers for its options (for example 16 means
"save as .docx"). The constants below give those numbers readable names,
so the engine code says FORMAT_DOCX instead of 16.

The values come from Word's own enumerations (WdSaveFormat,
WdExportFormat, WdBreakType and so on) in the Word VBA reference.
"""

# MailMerge.Destination: put the merge result in a new document.
SEND_TO_NEW_DOCUMENT = 0

# Document.Close: close without saving.
DO_NOT_SAVE_CHANGES = 0

# SaveAs2 FileFormat: Word .docx.
FORMAT_DOCX = 16

# ExportAsFixedFormat ExportFormat: PDF.
EXPORT_PDF = 17

# ExportAsFixedFormat OptimizeFor: print quality (full resolution).
OPTIMIZE_FOR_PRINT = 0

# InsertBreak: section break that starts on the next page.
SECTION_BREAK_NEXT_PAGE = 7

# MailMerge.MainDocumentType: form letters (one page per record).
FORM_LETTERS = 0

# AutomationSecurity: disable all macros. This also stops Word from asking
# to confirm the data source, which would pause an unattended run.
DISABLE_MACROS = 3


def load_word():
    """
    Import pywin32 and return (pythoncom, win32com.client).

    The import happens here, when Word is first needed, and not at the top
    of the file. That keeps the rest of DocuMate importable on machines
    without pywin32, for example when running the tests or the docxtpl
    versions.

    Raises RuntimeError with install instructions if pywin32 is missing.
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
