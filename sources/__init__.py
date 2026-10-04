"""
Sources: where records are read from, and where PRINTED is written back.

    excel.py      the Excel workbook                        v3, X, Y
    database.py   shared database logic                     Z, O
    postgres.py   PostgreSQL backend    (psycopg2)
    azuresql.py   Azure SQL backend     (pyodbc)

Every source has the same two methods:

    load_records()       returns a table (DataFrame) of records to print
    mark_printed(data)   sets STATUS to PRINTED for those records

Because every source returns the same shape of table, the rest of the
pipeline works the same whichever source a version uses.

Z and O use one of the two database backends, chosen by DOCUMATE_DB_BACKEND
in .env. The backend classes are not imported at the top of this file,
because importing one also imports its database driver. make_database_source()
below imports only the backend that is selected, so you only need that
backend's driver installed.
"""

from sources.excel import ExcelSource
from sources.database import DatabaseSource


def make_database_source():
    """
    Open a connection to the selected database and return its source object.

    The backend comes from DOCUMATE_DB_BACKEND in .env, not from the caller,
    so everything that talks to the database agrees on which one it is:
    versions/__init__.py builds Z and O from this, and tools/seed_database.py
    uses it too.

    The backend module is imported here rather than at the top of this file,
    so only the selected backend's driver (pyodbc or psycopg2) has to be
    installed, and the Excel versions need neither.

    The object it returns carries everything a caller needs:

        .connection  .cursor     the open connection
        .queries                 the SQL for this backend
        .sql(template)           fills in the {schema} placeholders
        .close()                 closes both handles
    """
    import importlib

    from setup import config

    backend = config.DATABASE_BACKENDS[config.database_backend()]
    module = importlib.import_module(backend["module"])
    source_class = getattr(module, backend["cls"])

    return source_class(config.database_settings(), config.database_schema())
