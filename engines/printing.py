"""
Sending the merged document to a printer.

This is the same job for every version. Both engines finish by writing one
merged .docx and closing Word behind them, so by the time anything is
printed there is no Word running and no difference between a file docxtpl
built and one Word's own Mail Merge built. Printing opens its own hidden
copy of Word, the way engines/pdf.py does, and prints the .docx rather
than the PDF - same renderer that produced it, and no PDF reader needed.

The document goes as one job, one copy. A hundred certificates are one
hundred pages of a single file, so the printer treats them as one stack
and pauses and resumes on its own if the paper runs out.

What "printed" means here
-------------------------
The job was accepted by the printer and spooled without raising an error.
That is not proof that paper exists: a job can still be cancelled at the
device, or jam, after this returns. It is the strongest thing that can be
known without asking somebody to go and count the stack.

So the printer is checked before the job is sent rather than after. An
offline or errored printer is caught while the run can still say so
usefully, instead of a job disappearing into a queue nobody is watching.
"""

import os

from setup import config
from setup import messages
from engines import word_com


def load_printing():
    """
    Import win32print, or explain why printing is unavailable.

    Separate from word_com.load_word() because printing can fail for its
    own reasons - pywin32 missing, not running on Windows - and the
    message should say which.
    """
    try:
        import win32print
    except ImportError as error:
        raise RuntimeError(messages.PRINTING_UNAVAILABLE.format(error=error))

    return win32print


def available_printers():
    """
    Every printer this PC can reach, and which one is the default.

    Returns (printers, default_name). Each printer is a dict with its name,
    its port, and whether that port means it writes a file instead of
    printing.

    Both locally attached and network printers are listed: Windows reports
    a USB printer and a \\\\server\\printer the same way once the connection
    exists, and Word takes the same name for either.
    """
    win32print = load_printing()

    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = []

    for entry in win32print.EnumPrinters(flags, None, 2):
        port = str(entry.get("pPortName") or "")
        printers.append({
            "name": entry.get("pPrinterName"),
            "port": port,
            "virtual": port.strip().upper() in config.VIRTUAL_PRINTER_PORTS,
        })

    try:
        default = win32print.GetDefaultPrinter()
    except Exception:
        # A machine with no printers at all has no default either.
        default = None

    return printers, default


def printer_problem(name):
    """
    What is wrong with this printer, or None when it looks ready.

    Checked before the job is sent. A printer that is offline or out of
    paper is worth saying so about now, while the run can still report it,
    rather than after a hundred pages have gone into a queue.
    """
    win32print = load_printing()

    handle = None
    try:
        handle = win32print.OpenPrinter(name)
        info = win32print.GetPrinter(handle, 2)
    except Exception as error:
        return messages.PRINTER_UNREACHABLE.format(printer=name, error=error)
    finally:
        if handle is not None:
            try:
                win32print.ClosePrinter(handle)
            except Exception:
                pass

    status = info.get("Status", 0)
    for flag, description in [
        (win32print.PRINTER_STATUS_OFFLINE, "offline"),
        (win32print.PRINTER_STATUS_ERROR, "reporting an error"),
        (win32print.PRINTER_STATUS_PAPER_OUT, "out of paper"),
        (win32print.PRINTER_STATUS_PAPER_JAM, "jammed"),
        (win32print.PRINTER_STATUS_NOT_AVAILABLE, "not available"),
        (win32print.PRINTER_STATUS_NO_TONER, "out of toner"),
        (win32print.PRINTER_STATUS_DOOR_OPEN, "open"),
    ]:
        if status & flag:
            return messages.PRINTER_NOT_READY.format(
                printer=name, problem=description)

    return None


def print_document(path, printer):
    """
    Print the merged document, once, on the named printer.

    Raises RuntimeError if Word is unavailable or the job is refused.
    Returns nothing: there is no job number to hand back, because Word's
    PrintOut does not give one.
    """
    if not os.path.exists(path):
        raise RuntimeError(messages.PRINT_FILE_MISSING.format(path=path))

    pythoncom, win32com = word_com.load_word()

    word = None
    document = None
    original_printer = None

    try:
        pythoncom.CoInitialize()

        # A separate Word process, like the PDF export uses. Dispatch would
        # attach to a window the user already has open and print from that.
        word = win32com.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = word_com.DISABLE_MACROS

        # ActivePrinter is a setting that sticks. Changing it here would
        # quietly change which printer the user's own Word prints to
        # afterwards, so it goes back the way it was found.
        try:
            original_printer = word.ActivePrinter
        except Exception:
            original_printer = None

        word.ActivePrinter = printer

        document = word.Documents.Open(
            os.path.abspath(path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        # Background=False waits for the job to be handed to the spooler,
        # so a refusal is raised here instead of vanishing silently.
        document.PrintOut(Background=False, Copies=1)

    except Exception as error:
        raise RuntimeError(messages.PRINT_FAILED.format(
            printer=printer, error=error))

    finally:
        if document is not None:
            try:
                document.Close(word_com.DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if word is not None:
            if original_printer:
                try:
                    word.ActivePrinter = original_printer
                except Exception:
                    pass
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
