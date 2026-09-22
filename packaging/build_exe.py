"""
Build the DocuMate .exe files with PyInstaller.

Usage (from the project folder):

    python packaging\\build_exe.py          build every exe in EXES
    python packaging\\build_exe.py v3       build only DocuMate_v3.exe

Requirements:  pip install pyinstaller

Output
------
The exes are written to dist\\ in the project folder. Keep them there.
When an exe runs, config.project_root() goes one folder up from the exe to
find files\\ (data, templates, output) and the .env file. An exe moved
somewhere else will not find them.

Adding an exe for another version
---------------------------------
1. Copy packaging\\documate_x.py to e.g. documate_z.py and change the key
   passed to launch().
2. Add a line to EXES below, e.g.  "z": ("documate_z.py", "DocuMate_Z").
"""

import os
import sys

import PyInstaller.__main__


PACKAGING = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(PACKAGING)

# key -> (launcher script in packaging\, exe name)
EXES = {
    "v3": ("documate_v3.py", "DocuMate_v3"),
    "x": ("documate_x.py", "DocuMate_X"),
}


def build(key):
    """Build one exe from its entry in EXES."""
    script, name = EXES[key]
    print("\nBuilding " + name + ".exe ...\n")

    PyInstaller.__main__.run([
        os.path.join(PACKAGING, script),
        "--name", name,
        "--onefile",            # a single .exe, no folder of DLLs beside it
        "--console",            # keep the console: the "mark as PRINTED?" prompt is typed there
        "--noconfirm",
        "--clean",
        "--distpath", os.path.join(ROOT, "dist"),
        "--workpath", os.path.join(ROOT, "build"),
        "--specpath", os.path.join(ROOT, "build"),
        "--paths", ROOT,
        "--paths", PACKAGING,
        # python-docx, docxtpl and docxcompose load XML template files from
        # inside their packages. These flags bundle those files; without
        # them the exe fails when it saves the first document.
        "--collect-data", "docx",
        "--collect-data", "docxtpl",
        "--collect-data", "docxcompose",
        "--hidden-import", "openpyxl",
        "--hidden-import", "win32timezone",
        # Left out to keep the exe smaller. v3 and X read Excel, so they do
        # not need the database drivers (used only by Z and O).
        "--exclude-module", "pytest",
        "--exclude-module", "tkinter",
        "--exclude-module", "psycopg2",
        "--exclude-module", "pyodbc",
    ])


if __name__ == "__main__":
    keys = sys.argv[1:] or list(EXES)
    for key in keys:
        build(key)
    print("\nDone. Exes are in " + os.path.join(ROOT, "dist"))
