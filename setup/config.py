"""
Paths and settings for DocuMate.

Everything DocuMate needs to find is decided here: the project folder, the
config.env settings file, the files/ folders, and the database connection details.

DocuMate is not tied to any one kind of document. Each run it needs an
Excel workbook, a Word template and the column that identifies a row; when
they are not in config.env, setup/picker.py asks for them and puts the answers in
the environment under the same names, so everything here reads them the
same way either way.

Secrets such as the database password are read from config.env and never
written in the code.

Folder layout
-------------
    <project root>/
        main.py
        config.env
        files/
            data/        Excel workbooks (any folder works - this is where
                         the file picker opens)
            templates/   Word templates (likewise)
            output/      generated .docx and .pdf files
            last_run.json  last run's choices, used as the pickers' defaults
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


# The one settings file, in the project root.
SETTINGS_FILENAME = "config.env"


def env_file():
    """Return the path of config.env, or None if there isn't one."""
    path = os.path.join(project_root(), SETTINGS_FILENAME)
    return path if os.path.isfile(path) else None


def load_environment():
    """
    Load config.env into the environment, if python-dotenv is installed.

    Variables already set in the environment are not overwritten. A line
    left blank (DOCUMATE_EXCEL_PATH=) counts as not set, so it is asked for
    or takes its default. Without python-dotenv, settings are read from the
    real environment only.
    """
    try:
        from dotenv import dotenv_values
    except ImportError:
        return

    path = env_file()
    if not path:
        return

    # utf-8-sig, because Notepad and PowerShell can save a file with an
    # invisible byte-order mark at the start. Read as plain utf-8 it sticks
    # to the first setting's name, and that setting - whatever is on line
    # one - is silently ignored.
    for name, value in dotenv_values(path, encoding="utf-8-sig").items():
        if value and value.strip() and name not in os.environ:
            os.environ[name] = value.strip()


def setting(name, default):
    """One setting: the text, or `default` when it is not set."""
    load_environment()
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip()


def required(name, what):
    """
    A setting that has no sensible default, or a clear error naming it.

    By the time a run needs one of these, the picker has normally asked for
    it, so this is reached by a run nobody could be asked in.
    """
    value = setting(name, "")
    if value:
        return value

    raise ValueError(
        name + " is not set.\n\n"
        "It is " + what + ".\n"
        "Run DocuMate and choose it when asked, or add it to config.env."
    )


def switch(name, default=True):
    """An on/off setting: 0 is off, anything else is on."""
    return setting(name, "1" if default else "0") != "0"


# ---------------------------------------------------------------------------
# Folders
# ---------------------------------------------------------------------------

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


def in_folder(folder, path):
    """A full path as it is; a bare file name or relative path inside `folder`."""
    path = os.path.expanduser(path)
    if os.path.isabs(path):
        return path
    return os.path.join(folder, path)


# ---------------------------------------------------------------------------
# This run's files  (chosen when DocuMate starts, or set in config.env)
# ---------------------------------------------------------------------------

def excel_path():
    """The workbook, DOCUMATE_EXCEL_PATH. A bare file name is looked for in files/data."""
    return in_folder(data_folder(), required(
        "DOCUMATE_EXCEL_PATH", "the Excel workbook to read"))


def sheet_name():
    """DOCUMATE_SHEET_NAME, or None for the first sheet in the workbook."""
    return setting("DOCUMATE_SHEET_NAME", "") or None


# Which setting holds the template for each engine. They are different
# files: a {{ placeholder }} template prints its braces under Mail Merge, and
# a Mail Merge template has no placeholders for docxtpl to fill.
TEMPLATE_SETTINGS = {
    "docxtpl": ("DOCUMATE_TEMPLATE",
                "the Word template with {{ placeholders }}, used by v3 and Z"),
    "mailmerge": ("DOCUMATE_TEMPLATE_MM",
                  "the Word Mail Merge template, used by X, Y and O"),
}


def template_path(engine):
    """The template for this engine. A bare file name is looked for in files/templates."""
    name, what = TEMPLATE_SETTINGS[engine]
    return in_folder(templates_folder(), required(name, what))


# Characters Windows will not accept in a file name.
UNSAFE_FILENAME = re.compile(r'[<>:"/\\|?*]+')


def output_name():
    """
    The middle of the output file name: DocuMateX_<this>_21102026.docx

    The template's own name, so FOREIGN_PCC_FORMAT.docx makes
    DocuMateX_FOREIGN_PCC_FORMAT_21102026.docx and two kinds of document
    never write files that look alike. DOCUMATE_OUTPUT_NAME replaces it.
    """
    name = setting("DOCUMATE_OUTPUT_NAME", "")
    if not name:
        for setting_name, _what in TEMPLATE_SETTINGS.values():
            template = setting(setting_name, "")
            if template:
                name = os.path.splitext(os.path.basename(template))[0]
                break
    name = UNSAFE_FILENAME.sub("_", name or "DOCUMENTS").replace(" ", "_")
    return name or "DOCUMENTS"


# ---------------------------------------------------------------------------
# How the rows are read
# ---------------------------------------------------------------------------

# plain            every column goes to the template under its own heading.
# birth_register   the birth register's raw entry sheet, converted to its
#                  certificate fields by sources/excel_format.py.
#
# Nobody has to choose: the picker recognises the birth register from its
# column headings and sets DOCUMATE_DATA_FORMAT for the run.
DATA_FORMATS = ["plain", "birth_register"]


def data_format():
    """plain or birth_register, from DOCUMATE_DATA_FORMAT."""
    name = setting("DOCUMATE_DATA_FORMAT", "plain").lower() or "plain"
    if name not in DATA_FORMATS:
        raise ValueError(
            "DOCUMATE_DATA_FORMAT is '" + name + "'. Use one of: "
            + ", ".join(DATA_FORMATS)
        )
    return name


def is_register():
    """True when the rows are the birth register's and need converting."""
    return data_format() == "birth_register"


# The column a birth register row is known by. Its raw sheet calls it S No.
REGISTER_ID_COLUMN = "Serial"


def id_column():
    """
    The column that tells one row from another, DOCUMATE_ID_COLUMN.

    Statuses are written back to the row with this value, so it must be
    filled in and different on every row. The birth register's is fixed.
    """
    if is_register():
        return REGISTER_ID_COLUMN
    return required(
        "DOCUMATE_ID_COLUMN",
        "the column that identifies each row, used to mark rows done")


# The birth register's own column rules. Its raw sheet is checked column by
# column, because several columns are joined into one certificate field and
# a gap inside the join would otherwise print as a stray comma.
#
# ADD 2 and ADD 3 are optional because an address does not always need
# three lines. ADD 1 is not: an address has to start somewhere, and a blank
# first line with a filled second one is a data-entry slip.
REGISTER_DATE_COLUMNS = ["Registration_date"]
REGISTER_OPTIONAL_COLUMNS = ["ADD 2", "ADD 3"]
REGISTER_DUPLICATE_COLUMN = "File_Number"

# The progress columns. DocuMate fills them in, so they may always be blank.
STATE_COLUMNS_OPTIONAL = ["STATUS", "Date_Printed", "Date_Issued"]


def date_columns():
    """
    Columns checked and printed as dd/mm/yyyy beyond the ones Excel already
    holds as dates - which are formatted whatever kind of document it is.
    A case file's DATES are added to these.
    """
    from setup import cases

    columns = list(REGISTER_DATE_COLUMNS) if is_register() else []
    return columns + [column for column in cases.dates() if column not in columns]


def optional_columns():
    """
    Columns the birth register allows to be blank. Any other sheet is
    checked only on the columns its template uses - see flow/checks.py.
    """
    extra = REGISTER_OPTIONAL_COLUMNS if is_register() else []
    return STATE_COLUMNS_OPTIONAL + extra


def duplicate_column():
    """The column checked for the same record entered twice."""
    return REGISTER_DUPLICATE_COLUMN if is_register() else id_column()


def sort_column():
    """The column documents are ordered by: the number in the ID column."""
    return id_column()


def file_number_prefix():
    """
    DOCUMATE_FILE_PREFIX: what every File_Number must start with.

    File numbers are issued with a fixed prefix, so one that does not start
    with it is either a typo or a record from somewhere else. Empty turns
    the check off; it only applies to a File_Number column.
    """
    return setting("DOCUMATE_FILE_PREFIX", "")


def mha_file_prefix():
    """DOCUMATE_MHA_PREFIX: the same check for an MHA_File_No column."""
    return setting("DOCUMATE_MHA_PREFIX", "")


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
    changed in config.env.

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
# pattern first to rule out SQL injection through config.env.
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
    Return the database connection details from config.env as a dict.

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
            "Database settings missing. Add these to your config.env file: "
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


def database_identifier(name, value):
    """
    A table or column name, checked the same way as the schema.

    Like the schema, these are written into the SQL text, so anything that
    is not a plain identifier is refused rather than run.
    """
    if value and not SCHEMA_NAME.match(value):
        raise ValueError(
            name + " is '" + value + "', which is not a plain name. Use "
            "letters, digits and underscores only - for example: student_id"
        )
    return value


def database_table():
    """
    The table Z and O read, DOCUMATE_DB_TABLE: every column of it, and each
    column becomes a template field.

    Empty means the birth register's own joined tables (sources/queries_*.py).
    """
    return database_identifier("DOCUMATE_DB_TABLE", setting("DOCUMATE_DB_TABLE", ""))


def database_id_column():
    """The column that identifies a row in DOCUMATE_DB_TABLE (default id)."""
    return database_identifier("DOCUMATE_DB_ID_COLUMN",
                               setting("DOCUMATE_DB_ID_COLUMN", "") or "id")


# ---------------------------------------------------------------------------
# Fixed rules
# ---------------------------------------------------------------------------

# How every date is written on a document, whatever shape it arrived in.
DATE_FORMAT = "%d/%m/%Y"

# Dates typed as text are read day-first, and only with these separators.
# "03/04/2025" is the 3rd of April, never the 4th of March - the parser is
# never allowed to guess, because a silent day/month swap is invisible on a
# finished document.
#
# "." is deliberately NOT here. Operators do type 8.11.26, and accepting it
# would read correctly - but it would also stop the report mentioning it, and
# the sheet would keep collecting dates in whichever shape each person prefers.
# Rejecting it puts the row in front of someone, who fixes the source. Add "."
# here if you would rather the tool took it quietly.
DATE_SEPARATORS = ["/", "-"]

# The status a record starts with, used when a sheet has no STATUS column
# yet. Anything that is not finished counts as pending, so the exact
# wording matters only to whoever reads the sheet.
PENDING_STATUS = "IN PROCESS"


# The status a record reaches once its document has been generated, but
# before anything has been sent to a printer. A run that generates and is
# not printed leaves records here, which says what actually happened.
COMPLETED_STATUS = "COMPLETED"

# The status that means a document has been sent to a printer.
DONE_STATUS = "PRINTED"

# Statuses that mean a record needs no more work. Anything else is pending,
# including a blank cell and a half-typed status, so a record is never
# skipped quietly. COMPLETED is in here because its document already
# exists - leaving it out would regenerate it on the next run.
FINISHED_STATUSES = [COMPLETED_STATUS, DONE_STATUS]


# Printers that write a file instead of putting ink on paper. They are
# recognised by their port: Windows gives Print to PDF and XPS the
# PORTPROMPT: port, OneNote nul:, and Fax SHRFAX:. Worth naming, because
# "Microsoft Print to PDF" is the default printer on a lot of machines and
# sending a hundred documents to it produces nothing and looks like it
# worked.
VIRTUAL_PRINTER_PORTS = ["PORTPROMPT:", "NUL:", "SHRFAX:"]

# Add the current year to every record as {{ year }}, so templates don't
# need it typed in. Set to False to turn it off.
INJECT_YEAR = True


def use_popups():
    """
    Show Windows popup messages. DOCUMATE_POPUPS=0 prints to the console
    instead. Popups are turned off automatically when not on Windows.
    """
    return switch("DOCUMATE_POPUPS")


def offer_printing():
    """
    Offer to print after generating. DOCUMATE_PRINTING=0 for an office that
    prints the merged file by hand, or for an unattended run where there is
    nobody to answer the question.
    """
    return switch("DOCUMATE_PRINTING")
