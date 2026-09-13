"""
Load records from a spreadsheet into PostgreSQL, and look at what's pending.

This was src/loadData.py. It isn't part of making documents - it's the
setup tool for the database that Z and O read from. It stays a separate
tool because it's the only thing here that writes new records.

    python -m tools.seed_database --insert
    python -m tools.seed_database --select
    python -m tools.seed_database --insert --file files/data/test_data/Test_Insert_data.xlsx

Connection details come from .env, same as everything else.
"""

import argparse
import os

import pandas as pd

from setup import config
from sources.queries import PENDING_APPLICANTS


DEFAULT_SHEET = "Sheet1"


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

INSERT_STATE = """
    INSERT INTO state (state_code, state_name)
    VALUES (%s, %s)
    ON CONFLICT (state_code) DO NOTHING;
"""

INSERT_AUTHORITY = """
    INSERT INTO ib_authority (
        ib_staff_authority_id,
        ib_staff_authority_name,
        ib_staff_authority_designation
    )
    VALUES (%s, %s, %s)
    ON CONFLICT (ib_staff_authority_id) DO NOTHING;
"""

INSERT_MHA = """
    INSERT INTO ministryofhomeaffairs (mha_file_number, mha_date)
    VALUES (%s, %s)
    ON CONFLICT (mha_file_number) DO NOTHING;
"""

INSERT_APPLICANT = """
    INSERT INTO applicant (
        file_number, name, sex, birth_date, place, state_code,
        name_of_father, name_of_mother,
        address_line_1, address_line_2, address_line_3,
        registration_date, mha_file_number, ib_staff_authority_id
    )
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);
"""

# Spreadsheet columns, in the order INSERT_APPLICANT wants them.
# The authority id is worked out separately and added on the end.
APPLICANT_COLUMNS = [
    "File_Number", "Name", "Sex", "Birth_Date", "Place", "State",
    "Name_Of_Father", "Name_Of_Mother",
    "Address_line_1", "Address_line_2", "Address_line_3",
    "Registration_Date", "MHA_FILE_NUMBER",
]


def insert_data(cursor, connection, workbook, sheet):
    """
    Load the spreadsheet into all four tables.

    Every insert says ON CONFLICT DO NOTHING, so running this twice is
    safe - rows that already exist are left alone, not duplicated.
    """
    data = pd.read_excel(workbook, sheet_name=sheet)
    data.columns = [str(c).strip() for c in data.columns]

    needed = APPLICANT_COLUMNS + ["MHA_DATE", "Signing_Authority"]
    missing = [c for c in needed if c not in data.columns]
    if missing:
        raise SystemExit(
            os.path.basename(workbook) + " is missing columns: " + ", ".join(missing)
        )

    # The lookup tables first - applicant rows point at both of them.
    for code in STATE_MASTER:
        cursor.execute(INSERT_STATE, (code.upper(), STATE_MASTER[code].upper()))

    for full_name in AUTHORITY_MAP:
        cursor.execute(INSERT_AUTHORITY, (
            AUTHORITY_MAP[full_name],
            full_name.split(",")[0].strip(),
            full_name.split(",")[-1].strip(),
        ))

    # MHA is the parent of applicant, so it has to go in first.
    mha_rows = data[["MHA_FILE_NUMBER", "MHA_DATE"]].drop_duplicates()
    for _, row in mha_rows.iterrows():
        cursor.execute(INSERT_MHA, (row["MHA_FILE_NUMBER"], row["MHA_DATE"]))

    # Now the applicants.
    unknown_authorities = set()
    inserted = 0

    for _, row in data.iterrows():
        authority = row["Signing_Authority"]
        authority_id = AUTHORITY_MAP.get(authority)

        if authority_id is None:
            # The original quietly put NULL here. We collect them and warn
            # instead, because NULL means a certificate with no signature.
            unknown_authorities.add(str(authority))

        values = [row[c] for c in APPLICANT_COLUMNS]
        values.append(authority_id)

        cursor.execute(INSERT_APPLICANT, tuple(values))
        inserted = inserted + 1

    connection.commit()
    print("DocuMate : inserted " + str(inserted) + " rows from " + os.path.basename(workbook))

    if unknown_authorities:
        print("\nDocuMate : these signing authorities are not in AUTHORITY_MAP,")
        print("           so those applicants have no authority set:")
        for name in sorted(unknown_authorities):
            print("             - " + name)


def select_data(cursor):
    """Print every applicant not yet marked PRINTED - what Z and O would read."""
    cursor.execute(PENDING_APPLICANTS)
    rows = cursor.fetchall()

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

    import psycopg2

    connection = psycopg2.connect(**config.database_settings())
    cursor = connection.cursor()

    try:
        if args.insert:
            workbook = args.file or default_workbook()
            insert_data(cursor, connection, workbook, args.sheet)
        else:
            select_data(cursor)

    except Exception:
        connection.rollback()
        raise

    finally:
        cursor.close()
        connection.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
