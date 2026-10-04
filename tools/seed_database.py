"""
Load records from a spreadsheet into the database, and look at what's pending.

This was src/loadData.py. It isn't part of making documents - it's the
setup tool for the database that Z and O read from. It stays a separate
tool because it's the only thing here that writes new records.

    python -m tools.seed_database --insert
    python -m tools.seed_database --select
    python -m tools.seed_database --insert --file files/data/test_data/Test_Insert_data.xlsx

Connection details come from .env, same as everything else, and so does the
database: this tool writes to whichever backend DOCUMATE_DB_BACKEND selects,
Azure SQL or PostgreSQL. Create the tables first with the matching script in
project/database_schema.

Two things differ between the backends, and both are handled here:

    parameter marks   PostgreSQL uses %s, pyodbc uses ?   (PLACEHOLDER)
    skipping rows     PostgreSQL has ON CONFLICT DO NOTHING and T-SQL has
                      nothing like it, so rows that are already there are
                      found first with existing_keys() and simply not sent

Because of the second one, --insert can be run as often as you like: rows
already in the database are left alone, never duplicated.
"""

import argparse
import os

import pandas as pd

from setup import config
from sources import make_database_source


DEFAULT_SHEET = "Sheet1"

# The parameter mark each driver expects, written as {p} in the SQL below.
# psycopg2 uses %s, pyodbc uses ?.
PLACEHOLDER = {
    "postgres": "%s",
    "azuresql": "?",
}


def default_workbook():
    return os.path.join(config.data_folder(), "test_data", "Test_Insert_data.xlsx")


# ---------------------------------------------------------------------------
# Reference data
# ---------------------------------------------------------------------------

# Signing authorities, keyed by the "NAME,DESIGNATION" text in the
# spreadsheet. We split it on the comma when inserting.
AUTHORITY_MAP = {
    "MONSHI G KATARI,VICE CONSUL": "MEA/IB/405/003",
    "GIRIRAJ SINGH KULSHEKHARA,VICE CONSUL AND ADMIN": "MEA/IB/405/004",
    "SHEESHACHELLAM SWAMINI, HOC AND CONSUL": "MEA/IB/405/002",
    "VEENA SAI RAJJAN KUMAR,CONSUL GENERAL": "MEA/IB/405/001",
}

STATE_MASTER = {
    "ACT": "Australian Capital Territory",
    "NSW": "New South Wales",
    "NT": "Northern Territory",
    "QLD": "Queensland",
    "SA": "South Australia",
    "TAS": "Tasmania",
    "VIC": "Victoria",
    "WA": "Western Australia",
}


# ---------------------------------------------------------------------------
# The inserts
# ---------------------------------------------------------------------------
#
# These are plain INSERTs that both databases accept. There is no
# ON CONFLICT DO NOTHING: T-SQL has no equivalent, so insert_data() checks
# what is already there first and skips those rows instead.
#
# {p} becomes the driver's parameter mark and {schema} becomes the schema
# prefix. sql_for() fills in both, in that order.

INSERT_STATE = """
    INSERT INTO {schema}state (state_code, state_name)
    VALUES ({p}, {p});
"""

INSERT_AUTHORITY = """
    INSERT INTO {schema}ib_authority (
        ib_staff_authority_id,
        ib_staff_authority_name,
        ib_staff_authority_designation
    )
    VALUES ({p}, {p}, {p});
"""

INSERT_MHA = """
    INSERT INTO {schema}ministryofhomeaffairs (mha_file_number, mha_date)
    VALUES ({p}, {p});
"""

INSERT_APPLICANT = """
    INSERT INTO {schema}applicant (
        file_number, name, sex, birth_date, place, state_code,
        name_of_father, name_of_mother,
        address_line_1, address_line_2, address_line_3,
        registration_date, mha_file_number, ib_staff_authority_id
    )
    VALUES ({p},{p},{p},{p},{p},{p},{p},{p},{p},{p},{p},{p},{p},{p});
"""

# Spreadsheet columns, in the order INSERT_APPLICANT wants them.
# The authority id is worked out separately and added on the end.
APPLICANT_COLUMNS = [
    "File_Number", "Name", "Sex", "Birth_Date", "Place", "State",
    "Name_Of_Father", "Name_Of_Mother",
    "Address_line_1", "Address_line_2", "Address_line_3",
    "Registration_Date", "MHA_FILE_NUMBER",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def sql_for(source, template, placeholder):
    """
    Turn one of the templates above into SQL this backend can run.

        {p}       the driver's parameter mark, %s or ?
        {schema}  the schema prefix, filled in by the source

    {p} is swapped with str.replace first, so {schema} is still there for
    source.sql() to fill in - the same trick queries_tsql.mark_printed_sql
    uses.
    """
    return source.sql(template.replace("{p}", placeholder))


def existing_keys(source, table, column):
    """
    Return the values already in one column, as a set of stripped strings.

    This is what replaces ON CONFLICT DO NOTHING. The table and column
    names come from the constants in this file, never from user input.

    The values are stripped because state_code is CHAR(3): a two-letter
    code comes back padded as "SA ", which would not match the "SA" in
    STATE_MASTER and would make us insert it a second time.
    """
    source.cursor.execute(source.sql("SELECT " + column + " FROM {schema}" + table + ";"))

    return set(str(row[0]).strip() for row in source.cursor.fetchall())


def native(value):
    """
    Turn a value read by pandas into a plain Python one the drivers accept.

        NaT / NaN          ->  None
        pandas Timestamp   ->  datetime.date   (the columns are all DATE)
        numpy int64        ->  int

    pyodbc is strict about this: it refuses a numpy integer outright, where
    psycopg2 happens to accept some of them. Converting here means the same
    spreadsheet loads into either database.
    """
    if pd.isna(value):
        return None

    if hasattr(value, "to_pydatetime"):
        return value.date()

    if hasattr(value, "item"):
        return value.item()

    return value


# ---------------------------------------------------------------------------
# The work
# ---------------------------------------------------------------------------

def insert_data(source, placeholder, workbook, sheet):
    """
    Load the spreadsheet into all four tables.

    Rows that are already in the database are skipped, so running this
    twice is safe - nothing is duplicated and nothing is overwritten.
    """
    data = pd.read_excel(workbook, sheet_name=sheet)
    data.columns = [str(c).strip() for c in data.columns]

    needed = APPLICANT_COLUMNS + ["MHA_DATE", "Signing_Authority"]
    missing = [c for c in needed if c not in data.columns]
    if missing:
        raise SystemExit(
            os.path.basename(workbook) + " is missing columns: " + ", ".join(missing)
        )

    cursor = source.cursor

    # The lookup tables first - applicant rows point at all of them.
    known_states = existing_keys(source, "state", "state_code")
    for code in STATE_MASTER:
        if code.upper() in known_states:
            continue
        cursor.execute(
            sql_for(source, INSERT_STATE, placeholder),
            (code.upper(), STATE_MASTER[code].upper()),
        )

    known_authorities = existing_keys(source, "ib_authority", "ib_staff_authority_id")
    for full_name in AUTHORITY_MAP:
        if AUTHORITY_MAP[full_name] in known_authorities:
            continue
        cursor.execute(sql_for(source, INSERT_AUTHORITY, placeholder), (
            AUTHORITY_MAP[full_name],
            full_name.split(",")[0].strip(),
            full_name.split(",")[-1].strip(),
        ))

    # MHA is the parent of applicant, so it has to go in first.
    known_mha = existing_keys(source, "ministryofhomeaffairs", "mha_file_number")
    mha_rows = data[["MHA_FILE_NUMBER", "MHA_DATE"]].drop_duplicates()
    for _, row in mha_rows.iterrows():
        if str(row["MHA_FILE_NUMBER"]).strip() in known_mha:
            continue
        cursor.execute(
            sql_for(source, INSERT_MHA, placeholder),
            (native(row["MHA_FILE_NUMBER"]), native(row["MHA_DATE"])),
        )

    # Now the applicants.
    known_applicants = existing_keys(source, "applicant", "file_number")
    unknown_authorities = set()
    inserted = 0
    skipped = 0

    for _, row in data.iterrows():
        if str(row["File_Number"]).strip() in known_applicants:
            skipped = skipped + 1
            continue

        authority = row["Signing_Authority"]
        authority_id = AUTHORITY_MAP.get(authority)

        if authority_id is None:
            # The original quietly put NULL here. We collect them and warn
            # instead, because NULL means a certificate with no signature.
            unknown_authorities.add(str(authority))

        values = [native(row[c]) for c in APPLICANT_COLUMNS]
        values.append(authority_id)

        cursor.execute(sql_for(source, INSERT_APPLICANT, placeholder), tuple(values))
        inserted = inserted + 1

    source.connection.commit()
    print("DocuMate : inserted " + str(inserted) + " rows from " + os.path.basename(workbook))

    if skipped:
        print("DocuMate : skipped " + str(skipped) + " rows already in the database.")

    if unknown_authorities:
        print("\nDocuMate : these signing authorities are not in AUTHORITY_MAP,")
        print("           so those applicants have no authority set:")
        for name in sorted(unknown_authorities):
            print("             - " + name)


def select_data(source):
    """Print every applicant not yet marked PRINTED - what Z and O would read."""
    # source.queries is the backend's own SQL module, so this is the very
    # same query Z and O run.
    source.cursor.execute(source.sql(source.queries.PENDING_APPLICANTS))
    rows = source.cursor.fetchall()

    if not rows:
        print("DocuMate : no pending records found.")
        return

    for row in rows:
        print(row)

    print("\nDocuMate : " + str(len(rows)) + " records found.")


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="tools.seed_database",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    what = parser.add_mutually_exclusive_group(required=True)
    what.add_argument("--insert", action="store_true", help="load the spreadsheet in")
    what.add_argument("--select", action="store_true", help="list pending applicants")

    parser.add_argument("--file", default=None, help="which spreadsheet to load")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="sheet name")

    args = parser.parse_args(argv)

    backend = config.database_backend()
    source = make_database_source()

    print("DocuMate : using " + source.label + ".")

    try:
        if args.insert:
            workbook = args.file or default_workbook()
            insert_data(source, PLACEHOLDER[backend], workbook, args.sheet)
        else:
            select_data(source)

    except Exception:
        source.connection.rollback()
        raise

    finally:
        source.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
