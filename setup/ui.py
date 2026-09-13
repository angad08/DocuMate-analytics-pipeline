"""
Popups and questions.

The old scripts called ctypes.windll.user32.MessageBoxW directly, in about
thirty different places, spread through the checks, the engines and run().
That meant Windows was baked into every part of the code and none of it
could be tested anywhere else.

Everything goes through here now. Not on Windows? It prints to the console
instead and the program carries on exactly the same.
"""

import sys

from setup import config

TITLE = "DocuMate"


def show_windows_popup(message):
    """Try to show a real Windows message box. Returns False if we can't."""
    if not config.USE_POPUPS:
        return False
    if not sys.platform.startswith("win"):
        return False

    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, message, TITLE, 0)
        return True
    except Exception:
        return False


def notify(message):
    """Show the user a message and wait for them to acknowledge it."""
    if not show_windows_popup(message):
        print("\n" + message + "\n")


def confirm(question):
    """
    Ask a yes/no question in the console.

    Same rule as the original: only "yes" or "y" counts as yes, ignoring
    case and spaces. Anything else is a no.
    """
    try:
        answer = input(question).strip().lower()
    except EOFError:
        return False

    return answer in ("yes", "y")


def progress(message):
    """Print over the same console line, for the record counter."""
    print(message, end="\r")
