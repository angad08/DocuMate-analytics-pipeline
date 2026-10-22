"""
Build the DocuMate .exe files with PyInstaller.

Usage (from the project folder):

    python packaging\\build_exe.py          build every exe in EXES
    python packaging\\build_exe.py v3       build only DocuMate_v3.exe
    python packaging\\build_exe.py z o      build the two database ones

Requirements:  pip install pyinstaller

Output
------
The exes are written to dist\\ in the project folder. Keep them there.
When an exe runs, config.project_root() goes one folder up from the exe to
find files\\ (data, templates, output) and the config.env file. An exe moved
somewhere else will not find them.

Each exe carries its own Python and every library it needs, so the PC it
runs on does not need Python installed. It still needs Microsoft Word for
the Mail Merge versions and for the PDF step, because those drive Word
itself rather than reimplementing it.

Adding an exe for another version
---------------------------------
1. Copy packaging\\documate_x.py to e.g. documate_p.py and change the key
   passed to launch().
2. Add a line to EXES below, e.g.  "p": ("documate_p.py", "DocuMate_P", EXCEL).
"""

import os
import sys

import PyInstaller.__main__


PACKAGING = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(PACKAGING)

# The window and taskbar icon. Built from logo\documate.ico.png by
# packaging\make_icon.py - Windows will not take a PNG here, whatever the
# file is called.
ICON = os.path.join(ROOT, "logo", "documate.ico")


# What a version needs bundled.
#
# The database drivers are imported by name at run time, inside a function,
# so that the Excel versions do not need them installed. PyInstaller cannot
# see an import like that, which is why Z and O name theirs here: without
# it the exe builds happily and then fails the moment it tries to connect.
#
# The Excel versions leave both out instead, which keeps their exe smaller.
EXCEL = {
    "hidden": [],
    "exclude": ["psycopg2", "pyodbc"],
}

DATABASE = {
    # The driver, and the module that wraps it. sources/__init__.py picks
    # the backend with importlib.import_module(), so neither the wrapper
    # nor its driver appears in any import statement PyInstaller can read.
    # Both have to be named, or the exe builds and then stops at
    # "No module named sources.postgres" the first time it is run.
    "hidden": [
        "psycopg2", "pyodbc",
        "sources.postgres", "sources.azuresql",
        "sources.queries_postgres", "sources.queries_tsql",
    ],
    "exclude": [],
}


# key -> (launcher script in packaging\, exe name, what it needs)
EXES = {
    "v3": ("documate_v3.py", "DocuMate_v3", EXCEL),
    "x": ("documate_x.py", "DocuMate_X", EXCEL),
    "y": ("documate_y.py", "DocuMate_Y", EXCEL),
    "z": ("documate_z.py", "DocuMate_Z", DATABASE),
    "o": ("documate_o.py", "DocuMate_O", DATABASE),
}


def build(key):
    """Build one exe from its entry in EXES."""
    script, name, needs = EXES[key]
    print("\nBuilding " + name + ".exe ...\n")

    options = [
        os.path.join(PACKAGING, script),
        "--name", name,
        "--onefile",            # a single .exe, no folder of DLLs beside it
        "--console",            # keep the console: the status and printer prompts are typed there
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
        # Printing enumerates printers through win32print. PyInstaller does
        # not see it, because it is imported inside a function so the Excel
        # versions can run without it.
        "--hidden-import", "win32print",
        "--exclude-module", "pytest",
        "--exclude-module", "tkinter",
    ]

    if os.path.exists(ICON):
        options += ["--icon", ICON]
    else:
        print("No icon at " + ICON + " - building without one.")

    for module in needs["hidden"]:
        options += ["--hidden-import", module]
    for module in needs["exclude"]:
        options += ["--exclude-module", module]

    PyInstaller.__main__.run(options)


if __name__ == "__main__":
    keys = sys.argv[1:] or list(EXES)

    unknown = [key for key in keys if key not in EXES]
    if unknown:
        raise SystemExit(
            "Unknown version(s): " + ", ".join(unknown)
            + ". Available: " + ", ".join(EXES)
        )

    for key in keys:
        build(key)

    print("\nDone. Exes are in " + os.path.join(ROOT, "dist"))
