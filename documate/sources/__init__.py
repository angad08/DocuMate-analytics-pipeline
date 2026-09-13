"""
Where records come from, and where PRINTED gets written back.

Two files, two sources:

    excel.py     -> used by v3, X and Y
    postgres.py  -> used by Z and O

Both do the same two jobs: load() hands back a table of records, and
mark_printed() writes the statuses back. Because they both hand back the
same shape, everything after them is shared.
"""

from documate.sources.excel import ExcelSource
from documate.sources.postgres import DatabaseSource
