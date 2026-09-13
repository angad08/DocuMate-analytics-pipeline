"""
Paths and settings.

Everything that used to sit hardcoded at the bottom of each script lives
here instead. Passwords and database hosts come from the .env file, never
from the code.
"""

import os
import sys


def project_root():
    """
    Find the project folder.

    Three cases:
      - DOCUMATE_ROOT is set  ->  use that (handy for testing)
      - running as a .exe     ->  the exe sits in dist\\, so go up one
      - running as a script   ->  this file is setup\\config.py,
                                  so go up one
    """
    override = os.getenv("DOCUMATE_ROOT")
    if override:
        return os.path.abspath(os.path.expanduser(override))

    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.dirname(sys.executable))

    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# The settings file. ".env" is the usual name, but Windows Explorer makes
# it awkward to create a file whose name starts with a dot, so anything
# ending in ".env" works too - DocuMate_env.env, env, and so on.
# ".env.example" is skipped on purpose: it is the blank template, and
# picking it up would look like the settings loaded when they hadn't.
ENV_FILENAMES = [".env", "DocuMate_env", "env.txt"]
ENV_SUFFIX = ".env"
ENV_TEMPLATE = ".env.example"


def env_file():
    """
    Find the settings file, or return None if there isn't one.

    Looks in the project root, alongside main.py. Exact names first, then any
    file ending in .env.
    """
    root = project_root()

    for name in ENV_FILENAMES:
        path = os.path.join(root, name)
        if os.path.exists(path):
            return path

    try:
        names = sorted(os.listdir(root))
    except OSError:
        return None

    for name in names:
        if name == ENV_TEMPLATE:
            continue
        if name.lower().endswith(ENV_SUFFIX):
            path = os.path.join(root, name)
            if os.path.isfile(path):
                return path

    return None


def load_environment():
    """Read the settings file if python-dotenv is installed. Optional."""
    try:
        from dotenv import load_dotenv
    except ImportError:
        return

    path = env_file()
    if path:
        load_dotenv(path, override=False)


def package_folder():
    """
    The project root. The package was flattened so all runtime folders sit
    directly beside main.py.
    """
    return project_root()


def files_folder():
    """
    What the code reads and writes.

    data, templates and output live in here, together and away from the
    code. If you ever move them back out to the project root, change this
    one function to use project_root() instead - it is the only place that
    location is decided.
    """
    return os.path.join(package_folder(), "files")


def data_folder():
    return os.path.join(files_folder(), "data")


def templates_folder():
    return os.path.join(files_folder(), "templates")


def output_folder():
    return os.path.join(files_folder(), "output")


# ---------------------------------------------------------------------------
# Excel settings  (used by v3, X, Y)
# ---------------------------------------------------------------------------

EXCEL_FILENAME = "DocuMate_DataFrame.xlsx"
SHEET_NAME = "DocuMateSRC"


def excel_path():
    load_environment()
    override = os.getenv("DOCUMATE_EXCEL_PATH")
    if override:
        return os.path.expanduser(override)
    return os.path.join(data_folder(), EXCEL_FILENAME)


def sheet_name():
    load_environment()
    return os.getenv("DOCUMATE_SHEET_NAME", SHEET_NAME)


# ---------------------------------------------------------------------------
# Database settings  (used by Z, O)
# ---------------------------------------------------------------------------

def database_settings():
    """
    Build the connection details for psycopg2, read from .env.

    If anything is missing we stop right here and name the variable,
    instead of failing later with a confusing connection error.
    """
    load_environment()

    settings = {
        "host": os.getenv("DOCUMATE_DB_HOST", ""),
        "database": os.getenv("DOCUMATE_DB_NAME", ""),
        "user": os.getenv("DOCUMATE_DB_USER", ""),
        "password": os.getenv("DOCUMATE_DB_PASSWORD", ""),
    }

    env_names = {
        "host": "DOCUMATE_DB_HOST",
        "database": "DOCUMATE_DB_NAME",
        "user": "DOCUMATE_DB_USER",
        "password": "DOCUMATE_DB_PASSWORD",
    }

    missing = []
    for key in settings:
        if not settings[key].strip():
            missing.append(env_names[key])

    if missing:
        raise ValueError(
            "Database settings missing. Add these to your .env file: "
            + ", ".join(missing)
        )

    settings["port"] = int(os.getenv("DOCUMATE_DB_PORT", "5432"))
    settings["sslmode"] = os.getenv("DOCUMATE_DB_SSLMODE", "require")
    return settings


# ---------------------------------------------------------------------------
# Switches
# ---------------------------------------------------------------------------

# Columns allowed to be empty. Every other column must have a value.
OPTIONAL_COLUMNS = ["STATUS", "Date_Printed", "Date_Issued"]

# The column used to spot the same applicant entered twice.
DUPLICATE_KEY = "File_Number"

# The column we sort by, using the number in it ("123/2025" sorts as 123).
SORT_COLUMN = "Serial"

# The status meaning "already done". Anything else counts as still pending.
DONE_STATUS = "PRINTED"

# Add the current year to every record so templates don't hardcode it.
# v3 did this, the others didn't. Set to False to turn it off.
INJECT_YEAR = True

# Show Windows popup boxes. Turns itself off when not on Windows.
USE_POPUPS = os.getenv("DOCUMATE_POPUPS", "1") != "0"
