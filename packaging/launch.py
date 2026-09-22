"""
Shared start-up code for the .exe launchers.

An exe is started by double-clicking, so there is no command line to type
the version into. Each launcher (documate_v3.py, documate_x.py) calls
launch() with its version key, which runs main.main() exactly as

    python main.py v3

would. All other settings (PDF on or off, polling, batch size) still come
from versions/registry.py.
"""

import multiprocessing
import os
import sys

# Add the project folder to the import path, so "import main" works when
# this file is run as a plain script. Inside a built exe the project is
# already bundled and this line has no effect.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import main


def launch(version):
    """Run DocuMate for one version, then wait for Enter before closing."""

    # Required in a frozen exe before starting worker processes (v3 fills
    # records in parallel). Does nothing when run as a normal script.
    multiprocessing.freeze_support()

    try:
        main.main([version])
    finally:
        # A double-clicked exe closes its window as soon as it finishes.
        # Waiting for Enter keeps the messages on screen to read.
        try:
            input("\nDocuMate : Press Enter to close...")
        except EOFError:
            pass
