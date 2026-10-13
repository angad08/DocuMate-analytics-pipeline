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


# How wide the bar itself is, in characters.
BAR_WIDTH = 24

# Solid blocks look better, but they are not printable everywhere: a Windows
# console on cp1252 raises UnicodeEncodeError on U+2588, which would stop a run
# partway through rendering. So ask the console what it can take rather than
# assuming - blocks on a UTF-8 terminal, hashes on an older one.
_characters = None


def bar_characters():
    """(filled, empty) - blocks where the console can print them, else ASCII."""
    global _characters
    if _characters is None:
        try:
            "█░".encode(sys.stdout.encoding or "ascii")
            _characters = ("█", "░")
        except (UnicodeEncodeError, LookupError, AttributeError):
            _characters = ("#", ".")
    return _characters

# Longest line the bar has printed so far, so a shorter one afterwards can be
# padded over it. Without this, going from "1440/1440" to "9/1440" would leave
# the tail of the longer line on screen.
_printed_width = 0


def progress_bar(done, total, template):
    """
    Draw a progress bar over the same console line.

    `template` comes from setup.messages and takes bar, percent, done and
    total, so the wording stays with every other piece of text DocuMate shows.

    Redrawing in place needs a real terminal. When output is piped to a file or
    captured, carriage returns pile thousands of half-lines into the log, so
    there it prints one ordinary line at each 10% instead, and always the last.
    """
    global _printed_width

    if total <= 0:
        return

    done = max(0, min(done, total))
    filled = round(BAR_WIDTH * done / total)
    block, blank = bar_characters()
    line = template.format(
        bar=block * filled + blank * (BAR_WIDTH - filled),
        percent=done * 100 // total,
        done=done,
        total=total,
    )

    if not sys.stdout.isatty():
        step = max(1, total // 10)
        if done == total or done % step == 0:
            print(line, flush=True)
        return

    print(line.ljust(_printed_width), end="\r", flush=True)
    _printed_width = max(_printed_width, len(line))

    if done == total:
        print()                 # leave the finished bar on its own line
        _printed_width = 0
