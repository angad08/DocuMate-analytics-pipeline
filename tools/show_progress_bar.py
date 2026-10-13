"""
Watch the progress bar run, without generating any documents.

Feeds setup.ui.progress_bar the same way the engines do, over a pretend 1,440
records, so the bar can be checked on a given console - including whether it
draws solid blocks or falls back to ASCII - without starting Word or touching
the workbook.

    python tools/show_progress_bar.py [records] [seconds]
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from setup import messages            # noqa: E402
from setup import ui                  # noqa: E402

records = int(sys.argv[1]) if len(sys.argv) > 1 else 1440
seconds = float(sys.argv[2]) if len(sys.argv) > 2 else 6.0

block, blank = ui.bar_characters()
print(f"console encoding : {sys.stdout.encoding}")
print(f"bar characters   : '{block}' filled, '{blank}' empty"
      f"{'  (solid blocks)' if block != '#' else '  (ASCII fallback)'}")
print(f"a terminal       : {sys.stdout.isatty()}")
print()

for done in range(1, records + 1):
    ui.progress_bar(done, records, messages.RECORD_PROGRESS)
    time.sleep(seconds / records)

print("done - no documents were made.")
