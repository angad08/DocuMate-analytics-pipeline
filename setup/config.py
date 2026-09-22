"""
Paths and settings for DocuMate.

Everything DocuMate needs to find is decided here: the project folder, the
.env settings file, the files/ folders, the Excel workbook, and the database
connection details.

Secrets such as the database password are read from .env (see .env.example)
and never written in the code.

Folder layout
-------------
    <project root>/
        main.py
        .env
        files/
            data/        the Excel workbook
            templates/   the Word templates
            output/      generated .docx and .pdf files
"""

import os
import re
import sys


def project_root():
    """
    Return the project folder.

      - DOCUMATE_ROOT is set  ->  that folder (useful for testing)
      - running as an .exe    ->  the folder above dist\\, where the exe sits
      - running as a script   ->  the folder above setup\\, where this file is
    """
    override = os.getenv("DOCUMATE_ROOT")
    if override:
        return os.path.abspath(os.path.expanduser(override))

    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.dirname(sys.executable))

    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# Names accepted for the settings file, checked in the project root.
# ".env" is the usual name. Windows Explorer makes it awkward to create a
# file starting with a dot, so a few other names work too, as does any file
# ending in ".env" (e.g. DocuMate_env.env).
# ".env.example" is always skipped: it is the blank template.
ENV_FILENAMES = [".env", "DocuMate_env", "env.txt"]
ENV_SUFFIX = ".env"
ENV_TEMPLATE = ".env.example"


def env_file():
    """
    Return the path of the settings file, or None if there isn't one.

    Checks ENV_FILENAMES in order first, then any other file ending in .env.
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
    """
    Load the settings file into the environment, if python-dotenv is installed.

    Variables already set in the environment are not overwritten. Without
    python-dotenv, settings are read from the real environment only.
    """
    try:
        from dotenv import load_dotenv
    except ImportError:
        return

    path = env_file()
    if path:
        load_dotenv(path, override=False)


def package_folder():
    """The folder holding main.py and files/. Currently the project root."""
    return project_root()


def files_folder():
    """
    The files/ folder, holding data/, templates/ and output/.

    This is the only place the location of files/ is set. Change it here
    to move all three folders at once.
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

# Default workbook (in files/data) and sheet. Override with
# DOCUMATE_EXCEL_PATH and DOCUMATE_SHEET_NAME in .env.
EXCEL_FILENAME = "DocuMate_DataFrame.xlsx"
SHEET_NAME = "DocuMateSRC"


def excel_path():
    """Full path to the workbook: DOCUMATE_EXCEL_PATH, or EXCEL_FILENAME in files/data."""
    load_environment()
    override = os.getenv("DOCUMATE_EXCEL_PATH")
    if override:
        return os.path.expanduser(override)
    return os.path.join(data_folder(), EXCEL_FILENAME)


def sheet_name():
    """The sheet to read: DOCUMATE_SHEET_NAME, or SHEET_NAME."""
    load_environment()
    return os.getenv("DOCUMATE_SHEET_NAME", SHEET_NAME)


# ---------------------------------------------------------------------------
# Database settings  (used by Z, O)
# ---------------------------------------------------------------------------

# The database backends Z and O can use.
#   port     default port, used when DOCUMATE_DB_PORT is not set
#   module   the module holding the source class
#   cls      the source class to create
DATABASE_BACKENDS = {
    "azuresql": {"port": "1433", "module": "sources.azuresql", "cls": "AzureSqlSource"},
    "postgres": {"port": "5432", "module": "sources.postgres", "cls": "PostgresSource"},
}

DEFAULT_BACKEND = "azuresql"


def database_backend():
    """
    Return the selected backend, "azuresql" or "postgres", from DOCUMATE_DB_BACKEND.

    The other connection settings (host, name, user, password) have the
    same names for both backends, so switching only needs this one value
    changed in .env.

    Raises ValueError for an unknown name.
    """
    load_environment()

    name = os.getenv("DOCUMATE_DB_BACKEND", DEFAULT_BACKEND).strip().lower()

    if name not in DATABASE_BACKENDS:
        raise ValueError(
            "DOCUMATE_DB_BACKEND is '" + name + "', which is not a database "
            "DocuMate knows. Use one of: " + ", ".join(sorted(DATABASE_BACKENDS))
        )

    return name


# Allowed schema names: letters, digits and underscores, not starting with
# a digit. The schema is written into the SQL text (databases do not accept
# a schema name as a query parameter), so it is checked against this
# pattern first to rule out SQL injection through .env.
SCHEMA_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def database_schema():
    """
    Return the schema holding the DocuMate tables, from DOCUMATE_DB_SCHEMA.

    Returns "" when it is not set, which means the connection's default
    schema: dbo on Azure SQL, public on PostgreSQL.

    Raises ValueError if the name is not a plain identifier.
    """
    load_environment()

    name = os.getenv("DOCUMATE_DB_SCHEMA", "").strip()

    if not name:
        return ""

    if not SCHEMA_NAME.match(name):
        raise ValueError(
            "DOCUMATE_DB_SCHEMA is '" + name + "', which is not a plain "
            "schema name. Use letters, digits and underscores only, starting "
            "with a letter or underscore - for example: documate"
        )

    return name


def database_settings():
    """
    Return the database connection details from .env as a dict.

        host, database, user, password   required
        port                             DOCUMATE_DB_PORT, or the backend's default
        sslmode                          PostgreSQL only, default "require"

    The keys match psycopg2.connect()'s keyword arguments. The Azure SQL
    backend turns the same dict into an ODBC connection string.

    Raises ValueError naming any required setting that is missing.
    """
    load_environment()
    backend = database_backend()

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

    settings["port"] = int(
        os.getenv("DOCUMATE_DB_PORT", DATABASE_BACKENDS[backend]["port"])
    )

    # sslmode is a psycopg2 option, so it is only added for PostgreSQL.
    # Azure SQL sets encryption in its connection string instead.
    if backend == "postgres":
        settings["sslmode"] = os.getenv("DOCUMATE_DB_SSLMODE", "require")

    return settings


# ---------------------------------------------------------------------------
# Switches
# ---------------------------------------------------------------------------

# Columns allowed to be empty. Every other column must have a value.
OPTIONAL_COLUMNS = ["STATUS", "Date_Printed", "Date_Issued"]

# The column used to find the same applicant entered twice.
DUPLICATE_KEY = "File_Number"

# The column records are sorted by, using its number ("123/2025" sorts as 123).
SORT_COLUMN = "Serial"

# The status that means a record is done. Any other value counts as pending.
DONE_STATUS = "PRINTED"

# Add the current year to every record, so templates don't need it typed in.
# Set to False to turn it off.
INJECT_YEAR = True

# Show Windows popup messages. Set DOCUMATE_POPUPS=0 to print to the console
# instead. Popups are turned off automatically when not on Windows.
USE_POPUPS = os.getenv("DOCUMATE_POPUPS", "1") != "0"
