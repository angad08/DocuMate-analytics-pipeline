"""
PostgreSQL backend. Used by Z and O when DOCUMATE_DB_BACKEND=postgres.

The shared reading and writing logic is in sources/database.py. This file
only covers what is specific to PostgreSQL:

    connecting          psycopg2.connect(**settings)
    the query timeout   SET statement_timeout = 0

Works with any PostgreSQL server, for example Supabase, Azure Database for
PostgreSQL, or a local install.

Requirements:  pip install psycopg2-binary
"""

from sources import queries_postgres
from sources.database import COLUMN_ORDER
from sources.database import TEMPLATE_FIELDS
from sources.database import DatabaseSource
from sources.database import as_date
from sources.database import build_record


# build_record and the column lists live in sources/database.py. They are
# also exported from here so "from sources.postgres import build_record"
# works, which the tests and the docs use.
__all__ = [
    "PostgresSource",
    "build_record",
    "as_date",
    "COLUMN_ORDER",
    "TEMPLATE_FIELDS",
]


class PostgresSource(DatabaseSource):
    """Reads pending applicants from PostgreSQL and writes statuses back."""

    label = "PostgreSQL"
    queries = queries_postgres

    def connect(self, settings):
        # Imported here, not at the top, so psycopg2 is only needed when the
        # PostgreSQL backend is actually used.
        import psycopg2

        # settings already has the keyword names psycopg2 expects:
        # host, database, user, password, port, sslmode.
        return psycopg2.connect(**settings)

    def remove_query_timeout(self):
        """Set PostgreSQL's statement_timeout to 0, which means no limit."""
        self.cursor.execute("SET statement_timeout = 0;")
