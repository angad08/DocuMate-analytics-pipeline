"""
Build the DocuMate .exe files with PyInstaller.

Usage (from the project folder):

    python packaging\\build_exe.py          build every exe in EXES
    python packaging\\build_exe.py v3       build only DocuMate_v3.exe
    python packaging\\build_exe.py z o      build the two database exes

Requirements:  pip install pyinstaller

This must be run on Windows. PyInstaller does not cross-compile: a Windows
.exe can only be built on Windows, by the same Python that runs DocuMate.

Output
------
The exes are written to dist\\ in the project folder. Keep them there.
When an exe runs, config.project_root() goes one folder up from the exe to
find files\\ (data, templates, output) and the .env file. An exe moved
somewhere else will not find them, so the exe is self-contained in its code
but not in its data - ship the project folder, not the .exe on its own.

Excel and database versions need different things bundled
---------------------------------------------------------
v3, X and Y read Excel and never import a database driver, so psycopg2 and
pyodbc are excluded to keep those exes smaller. Z and O read a database and
must have the driver bundled. Both drivers are imported inside functions
rather than at the top of a module, so PyInstaller cannot see them by
following imports - they are named as hidden imports instead.

Adding an exe for another version
---------------------------------
1. Copy packaging\\documate_z.py to e.g. documate_w.py and change the key
   passed to launch().
2. Add a line to EXES below, e.g.
   "w": ("documate_w.py", "DocuMate_W", DATABASE).
"""

import os
import sys

# PyInstaller is imported inside build() rather than here, so that running
# this file on a non-Windows machine gives the platform message below
# instead of a confusing ImportError.

PACKAGING = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(PACKAGING)
ICON = os.path.join(ROOT, "logo", "documate.ico")

# Which source a version reads. It decides whether the database drivers are
# bundled or excluded.
EXCEL = "excel"
DATABASE = "database"

# key -> (launcher script in packaging\, exe name, source)
EXES = {
    "v3": ("documate_v3.py", "DocuMate_v3", EXCEL),
    "x": ("documate_x.py", "DocuMate_X", EXCEL),
    "y": ("documate_y.py", "DocuMate_Y", EXCEL),
    "z": ("documate_z.py", "DocuMate_Z", DATABASE),
    "o": ("documate_o.py", "DocuMate_O", DATABASE),
}


def build(key):
    """Build one exe from its entry in EXES."""
    import PyInstaller.__main__

    script, name, source = EXES[key]
    print("\nBuilding " + name + ".exe ...\n")

    options = [
        os.path.join(PACKAGING, script),
        "--name", name,
        "--onefile",            # a single .exe, no folder of DLLs beside it
        "--console",            # keep the console: the prompts are typed there
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
        # Printing uses win32print, reached through engines/printing.py.
        "--hidden-import", "win32print",
        "--exclude-module", "pytest",
        "--exclude-module", "tkinter",
    ]

    # The icon is optional: a missing .ico should not stop a build.
    if os.path.exists(ICON):
        options += ["--icon", ICON]
    else:
        print("Note: " + ICON + " not found - building without an icon.")

    if source == DATABASE:
        # Imported inside functions in sources/postgres.py and
        # sources/azuresql.py, so they have to be named explicitly.
        options += [
            "--hidden-import", "psycopg2",
            "--hidden-import", "pyodbc",
        ]
    else:
        options += [
            "--exclude-module", "psycopg2",
            "--exclude-module", "pyodbc",
        ]

    PyInstaller.__main__.run(options)


if __name__ == "__main__":
    if os.name != "nt":
        raise SystemExit(
            "build_exe.py must be run on Windows. PyInstaller does not "
            "cross-compile, so a Windows .exe cannot be built on "
            + sys.platform + "."
        )

    keys = sys.argv[1:] or list(EXES)
    for key in keys:
        if key not in EXES:
            raise SystemExit(
                "Unknown version '" + key + "'. Available: "
                + ", ".join(EXES)
            )
    for key in keys:
        build(key)
    print("\nDone. Exes are in " + os.path.join(ROOT, "dist"))
