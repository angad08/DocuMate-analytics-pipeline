"""
The database data source. Shared by Z and O.

Z and O had exactly the same load_data() and update_sql_status() code,
word for word - the only real difference between those two files was the
engine. This is that code, once.
"""

import datetime

import pandas as pd

from setup import messages
from sources.queries import ADD_DATE_ISSUED_COLUMN
from sources.queries import DATE_ISSUED_EXISTS
from sources.queries import MARK_PRINTED
from sources.queries import PENDING_APPLICANTS


# The columns the query returns, in order. Kept next to the query so a
# change to one is obviously a change to the other.
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

# The field names the Word templates use.
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
    """Show a date as dd/mm/yyyy. Copes with it already being text."""
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    return str(value)


def build_record(row, year):
    """
    Turn one row from the database into what a Word template expects.

    Serial becomes "123/2025" and dates become dd/mm/yyyy here, not in the
    SQL, so the database keeps proper date types and only the presentation
    happens in Python.
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
    """Reads pending applicants, and writes statuses back."""

    # Shown in messages, e.g. "Reading PostgreSQL..."
    label = "PostgreSQL"

    # How many rows per UPDATE, so a big run doesn't time out.
    update_batch_size = 300

    def __init__(self, settings):
        # Imported here so the Excel versions don't need the driver installed.
        import psycopg2

        self.connection = psycopg2.connect(**settings)
        self.cursor = self.connection.cursor()

    # -- reading ------------------------------------------------------------

    def validate_source(self):
        """Make sure the connection actually works before we start."""
        self.cursor.execute("SELECT 1;")
        self.cursor.fetchone()

    def load_records(self):
        """
        Read the pending applicants into a table.

        The filtering already happened in the SQL, which is why Z and O run
        with checks turned off in the version list - there is nothing left
        for the checks to filter.
        """
        self.cursor.execute(PENDING_APPLICANTS)
        rows = self.cursor.fetchall()

        year = datetime.datetime.now().year
        records = []
        for row in rows:
            records.append(build_record(row, year))

        # Passing the column names keeps the table the same shape even when
        # there are no rows, so later code behaves the same either way.
        return pd.DataFrame(records, columns=TEMPLATE_FIELDS)

    def close(self):
        """Close the cursor and connection. Safe to call twice."""
        for handle in [self.cursor, self.connection]:
            try:
                handle.close()
            except Exception:
                pass

    # -- writing back -------------------------------------------------------

    def mark_printed(self, data):
        """
        Mark exactly the records from this run as PRINTED.

        Unlike the Excel version this targets specific Serial numbers, so a
        record added while we were working isn't touched.
        """
        if data is None or data.empty:
            print(messages.NOTHING_TO_UPDATE)
            return

        try:
            # Big updates can hit the default timeout.
            self.cursor.execute("SET statement_timeout = 0;")

            # Safety net for a database set up without date_issued.
            self.cursor.execute(DATE_ISSUED_EXISTS)
            if not self.cursor.fetchone():
                self.cursor.execute(ADD_DATE_ISSUED_COLUMN)
                self.connection.commit()

            # "123/2025" -> 123
            serials = []
            for value in data["Serial"]:
                serials.append(int(str(value).split("/")[0]))

            today = datetime.date.today()
            updated = 0

            for start in range(0, len(serials), self.update_batch_size):
                batch = serials[start:start + self.update_batch_size]
                self.cursor.execute(MARK_PRINTED, (today, batch))
                updated = updated + self.cursor.rowcount
                self.connection.commit()

            print(messages.DB_UPDATED.format(count=updated))

        except Exception as error:
            # Undo everything so we don't leave a half-finished update.
            self.connection.rollback()
            print(messages.DB_UPDATE_FAILED.format(error=error))
