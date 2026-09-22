"""
Database source, shared by both backends. Used by Z and O.

This file holds everything that is the same for PostgreSQL and Azure SQL:
reading the pending applicants, turning each row into a record for the
template, and marking records as PRINTED afterwards.

A backend only supplies the parts that differ between the two databases:

    how to connect           connect()
    how to lift the timeout  remove_query_timeout()
    the SQL itself           a queries module

    sources/postgres.py   psycopg2  + sources/queries_postgres.py
    sources/azuresql.py   pyodbc    + sources/queries_tsql.py

Which backend runs is set by DOCUMATE_DB_BACKEND in .env (read in
setup/config.py, used in versions/__init__.py).

Flow
----
1. __init__ connects to the database.
2. load_records() runs PENDING_APPLICANTS (every applicant not yet PRINTED)
   and turns each row into a record with build_record().
3. The pipeline makes the Word file from those records.
4. mark_printed() sets STATUS = 'PRINTED' and date_issued = today for
   exactly those Serial numbers, in batches.
"""

import datetime

import pandas as pd

from setup import messages


# The columns PENDING_APPLICANTS returns, in the same order as its SELECT.
# If you change the query, update this list to match.
COLUMN_ORDER = [
    "file_number",
    "serial",
    "name",
    "sex",
    "birth_date",
    "place_of_birth",
    "name_of_father",
    "name_of_mother",
    "address",
    "registration_date",
    "mha_file_number",
    "mha_date",
    "signing_authority_name",
    "authority_name_designation",
]

# The merge field names used in the Word templates.
TEMPLATE_FIELDS = [
    "File_Number",
    "Serial",
    "Name",
    "Sex",
    "When_and_where_born",
    "Name_of_the_Father",
    "Name_of_the_Mother",
    "Description_and_residence_of_informant",
    "Registration_date",
    "MHA_File_And_date",
    "Signing_Authority_Name",
    "Signing_Authority_Name_Designation",
]


def as_date(value):
    """
    Format a date as dd/mm/yyyy. Text is returned unchanged.

        as_date(date(2025, 3, 7))  ->  "07/03/2025"
    """
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    return str(value)


def build_record(row, year):
    """
    Turn one database row into the fields a Word template expects.

    row    one row from PENDING_APPLICANTS, in COLUMN_ORDER
    year   added to the serial number, e.g. 123 -> "123/2025"

    Returns a dict keyed by TEMPLATE_FIELDS. Some fields combine columns,
    for example When_and_where_born is "<birth date>, <place of birth>".

    Dates are formatted here rather than in the SQL, so the database keeps
    real date types.
    """
    r = dict(zip(COLUMN_ORDER, row))

    return {
        "File_Number": r["file_number"],
        "Serial": str(r["serial"]) + "/" + str(year),
        "Name": r["name"],
        "Sex": r["sex"],
        "When_and_where_born": as_date(r["birth_date"]) + ", " + str(r["place_of_birth"]),
        "Name_of_the_Father": r["name_of_father"],
        "Name_of_the_Mother": r["name_of_mother"],
        "Description_and_residence_of_informant": r["address"],
        "Registration_date": as_date(r["registration_date"]),
        "MHA_File_And_date": str(r["mha_file_number"]) + ", " + as_date(r["mha_date"]),
        "Signing_Authority_Name": r["signing_authority_name"],
        "Signing_Authority_Name_Designation": r["authority_name_designation"],
    }


class DatabaseSource:
    """
    Reads pending applicants and writes statuses back.

    This is a base class. Use PostgresSource or AzureSqlSource, which set
    the class attributes below and implement connect() and
    remove_query_timeout().

    settings   connection details from config.database_settings()
    schema     schema name from DOCUMATE_DB_SCHEMA; "" uses the default
    """

    # Name used in messages, e.g. "Reading PostgreSQL...".
    label = "database"

    # The queries module for this database. It must provide
    # PENDING_APPLICANTS, DATE_ISSUED_EXISTS, ADD_DATE_ISSUED_COLUMN,
    # mark_printed_sql() and mark_printed_params().
    queries = None

    # Rows per UPDATE when marking records PRINTED. Smaller batches avoid
    # timeouts on large runs. On Azure SQL this must also stay below
    # MAX_SERIALS_PER_UPDATE in queries_tsql.py.
    update_batch_size = 300

    def __init__(self, settings, schema=""):
        # An empty schema means the connection's default schema:
        # dbo on Azure SQL, public on PostgreSQL.
        self.schema = schema
        self.connection = self.connect(settings)
        self.cursor = self.connection.cursor()

    def sql(self, template):
        """
        Fill in the schema placeholders in a query.

        Every query runs through this method, so the schema is applied in
        one place.

            {schema}         "documate." when DOCUMATE_DB_SCHEMA=documate,
                             so "FROM {schema}applicant" becomes
                             "FROM documate.applicant"
            {schema_filter}  an extra "AND table_schema = '...'" condition
                             for the information_schema lookup

        With no schema set, both placeholders become empty strings.
        """
        if self.schema:
            prefix = self.schema + "."
            condition = " AND table_schema = '" + self.schema + "'"
        else:
            prefix = ""
            condition = ""

        return template.format(schema=prefix, schema_filter=condition)

    # -- implemented by each backend ----------------------------------------

    def connect(self, settings):
        """Open and return a database connection."""
        raise NotImplementedError

    def remove_query_timeout(self):
        """
        Turn off the query timeout so a large UPDATE can finish.

        PostgreSQL does this with a SQL statement; pyodbc does it with an
        attribute on the connection.
        """
        raise NotImplementedError

    # -- reading ------------------------------------------------------------

    def validate_source(self):
        """Run a trivial query to confirm the connection works."""
        self.cursor.execute("SELECT 1;")
        self.cursor.fetchone()

    def load_records(self):
        """
        Read the pending applicants and return them as a DataFrame.

        The SQL already filters out PRINTED records, which is why Z and O
        have check_records=False in versions/registry.py.
        """
        self.cursor.execute(self.sql(self.queries.PENDING_APPLICANTS))
        rows = self.cursor.fetchall()

        year = datetime.datetime.now().year
        records = []
        for row in rows:
            records.append(build_record(row, year))

        # Passing the column names gives the table the right columns even
        # when there are no rows.
        return pd.DataFrame(records, columns=TEMPLATE_FIELDS)

    def close(self):
        """Close the cursor and the connection. Safe to call more than once."""
        for handle in [self.cursor, self.connection]:
            try:
                handle.close()
            except Exception:
                pass

    # -- writing back -------------------------------------------------------

    def mark_printed(self, data):
        """
        Mark the records from this run as PRINTED, with today's date.

        data   the DataFrame that was merged; its Serial column says which
               applicants to update

        Only these exact Serial numbers are updated, so an applicant added
        to the database during the run is left pending for the next run.

        Each batch is committed on its own. If a batch fails, that batch is
        rolled back and the error is printed; batches committed before it
        stay PRINTED.
        """
        if data is None or data.empty:
            print(messages.NOTHING_TO_UPDATE)
            return

        try:
            self.remove_query_timeout()

            # Add the date_issued column if this database does not have it.
            self.cursor.execute(self.sql(self.queries.DATE_ISSUED_EXISTS))
            if not self.cursor.fetchone():
                self.cursor.execute(self.sql(self.queries.ADD_DATE_ISSUED_COLUMN))
                self.connection.commit()

            # Get the numbers back out of the serials: "123/2025" -> 123
            serials = []
            for value in data["Serial"]:
                serials.append(int(str(value).split("/")[0]))

            today = datetime.date.today()
            updated = 0

            # Update in batches of update_batch_size, committing each one.
            for start in range(0, len(serials), self.update_batch_size):
                batch = serials[start:start + self.update_batch_size]

                # The queries module builds the SQL and its parameters for
                # this batch size (see mark_printed_sql in each module).
                self.cursor.execute(
                    self.sql(self.queries.mark_printed_sql(len(batch))),
                    self.queries.mark_printed_params(today, batch),
                )

                updated = updated + self.cursor.rowcount
                self.connection.commit()

            print(messages.DB_UPDATED.format(count=updated))

        except Exception as error:
            # Undo any uncommitted changes from the failed batch.
            self.connection.rollback()
            print(messages.DB_UPDATE_FAILED.format(error=error))
