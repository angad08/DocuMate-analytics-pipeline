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
in .env. The backend classes are not imported here, because importing one
also imports its database driver. versions/__init__.py imports only the
backend that is selected, so you only need that backend's driver installed.
"""

from sources.excel import ExcelSource
from sources.database import DatabaseSource
