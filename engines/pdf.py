"""
PDF export for the merged Word file.

Used by both engines when a version has to_pdf=True in versions/registry.py.

How it works
------------
1. The engine saves the merged .docx as normal.
2. Word opens that .docx and exports it with ExportAsFixedFormat.
3. The PDF is saved next to the .docx, with the same name:

       files/output/DocuMateX_BIRTH_REGISTRATION_17092026.docx
       files/output/DocuMateX_BIRTH_REGISTRATION_17092026.pdf

The PDF is converted from the finished .docx instead of being built
separately. That way every record is filled only once, and the PDF always
matches the Word file page for page.

Two ways in
-----------
    export_document(document, docx_path)
        For a document Word already has open. The Mail Merge engine uses
        this, because Word is still running when the merge finishes.

    convert_to_pdf(docx_path)
        For a .docx saved on disk. Starts a hidden copy of Word, exports,
        then closes Word again. The docxtpl engine uses this, because it
        never opens Word itself.

If the export fails, the error is printed and the run carries on. The .docx
is already saved by then, so no output is lost.

Needs Windows, Microsoft Word and pywin32.
"""

import os
import time

from setup import messages
from engines import word_com


def pdf_path_for(docx_path):
    """
    Return the PDF path for a Word file: same folder, same name, .pdf.

        pdf_path_for("files/output/report.docx")  ->  ".../files/output/report.pdf"
    """
    return os.path.splitext(os.path.abspath(docx_path))[0] + ".pdf"


def export_document(document, docx_path):
    """
    Export a document that Word already has open.

    document    the open Word document (a COM object)
    docx_path   where the .docx was saved; the PDF goes beside it

    Returns the PDF path, or None if the export failed.
    """
    pdf_path = pdf_path_for(docx_path)
    print(messages.PDF_START)
    started = time.time()

    try:
        document.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=word_com.EXPORT_PDF,
            OpenAfterExport=False,
            OptimizeFor=word_com.OPTIMIZE_FOR_PRINT,
        )
    except Exception as error:
        print(messages.PDF_FAILED.format(error=error))
        return None

    print(messages.PDF_SAVED.format(path=pdf_path, seconds=time.time() - started))
    return pdf_path


def convert_to_pdf(docx_path):
    """
    Open a saved .docx in a hidden copy of Word and export it to PDF.

    Returns the PDF path, or None if Word is not available or the export
    failed. Word is always closed afterwards, even after an error.
    """
    try:
        pythoncom, win32com = word_com.load_word()
    except RuntimeError as error:
        print(messages.PDF_FAILED.format(error=error))
        return None

    word = None
    document = None

    try:
        pythoncom.CoInitialize()

        # DispatchEx starts a new, separate Word process. Dispatch would
        # attach to a Word window the user already has open instead.
        word = win32com.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = word_com.DISABLE_MACROS

        # Read-only: the export never changes the .docx.
        document = word.Documents.Open(
            os.path.abspath(docx_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        return export_document(document, docx_path)

    except Exception as error:
        print(messages.PDF_FAILED.format(error=error))
        return None

    finally:
        # Clean up in reverse order: document, then Word, then COM.
        if document is not None:
            try:
                document.Close(word_com.DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
