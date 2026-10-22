"""
Turning a raw entry sheet into the shape the certificates expect.

The workbook used to carry two sheets: the raw one people type into, and a
second one built from it with Excel formulas - UPPER(), TEXT(), a few
concatenations - which was the sheet DocuMate actually read. That second
sheet had to be built by hand before DocuMate could run, and it was easy to
break: rename it, re-save it the wrong way, or drag a formula one row short
and the run either failed or quietly skipped records.

The formulas were never complicated. They are here instead, so the raw sheet
is the only thing anyone has to maintain.

This is the same job sources/database.py does in build_record(): the database
stores normalised columns and assembles the certificate fields on the way
out. Excel was the odd one out in keeping that step in the spreadsheet.

Both shapes still work. A workbook that already has the prepared columns is
read as before - see looks_prepared() - so nothing that worked yesterday
stops working.
"""

import datetime

import pandas as pd

from setup import config
from setup import messages


# ---------------------------------------------------------------------------
# The two shapes
# ---------------------------------------------------------------------------

# What the raw entry sheet is expected to have, as the headers people type.
# The keys are what this module calls them internally; the values are what
# has to be in the sheet.
# Some columns answer to more than one name. "S No." is what the register
# uses, but "Serial" is the obvious thing to call it and is what the
# finished column is called, so both are accepted.
RAW_ALIASES = {
    "S No.": ["S No.", "Serial"],
    # "Sighning_Authority" is how the live register and its templates have
    # always spelt it.
    "Signing_Authority_Name": ["Signing_Authority_Name", "Signing_Authority",
                               "Signing Authority", "Sighning_Authority"],
}

RAW_COLUMNS = {
    "serial": "S No.",
    "name": "Name",
    "birth_date": "Date of birth",
    "place": "Place of birth",
    "file_number": "File No",
    "sex": "Sex",
    "father": "Name of Father",
    "mother": "Name of Mother",
    "registration_date": "Registration date",
    "address_1": "ADD 1",
    "address_2": "ADD 2",
    "address_3": "ADD 3",
    "mha_file_number": "MHA_File_No",
    "mha_date": "MHA_Date",
}

# What the certificates expect, and what the rest of DocuMate works with.
PREPARED_COLUMNS = [
    "Serial",
    "File_Number",
    "Name",
    "Sex",
    "Name_of_the_Father",
    "Name_of_the_Mother",
    "When_and_where_born",
    "Description_and_residence_of_informant",
    "Registration_date",
    "MHA_File_And_date",
]

# Columns that hold progress rather than applicant details. They are not
# derived from anything - DocuMate writes them - so they are carried across
# when the raw sheet has them and defaulted when it doesn't.
STATE_COLUMNS = ["STATUS", "Date_Printed", "Signing_Authority_Name", "Date_Issued"]

# Columns that print in capitals on the certificate. Upper-casing is done
# here rather than asked of whoever types the sheet: it is not a mistake to
# type a name in lower case, so there is nothing to flag - the certificate
# just needs it one way. The database versions do the same with UPPER() in
# their SQL, so both routes print alike.
UPPERCASE_COLUMNS = [
    "Name",
    "Name_of_the_Father",
    "Name_of_the_Mother",
    "Description_and_residence_of_informant",
    "Signing_Authority_Name",
    "File_Number",
]

# MHA file numbers are left alone. The real ones are plain numbers, so
# there are no letters to raise, and forcing a number through .upper()
# only risks changing how it is written for nothing.


# Three columns are enough to tell the shapes apart, and all three are ones
# the raw sheet could not have by accident.
PREPARED_MARKERS = ["Serial", "File_Number", "When_and_where_born"]


def looks_prepared(frame):
    """
    True when the sheet is already in certificate shape.

    Checked by column names rather than by sheet name, because the name is
    the part people change.
    """
    return all(column in frame.columns for column in PREPARED_MARKERS)


def is_register_sheet(frame):
    """
    True when this is the birth register - its raw entry sheet, or a sheet
    already in certificate shape - recognised by its column headings.

    "Most of the raw columns" counts, so that a register with one renamed
    heading is still treated as one and stopped with that heading named,
    rather than being read as some other kind of sheet.
    """
    return looks_prepared(frame) or raw_columns_found(frame) >= len(RAW_COLUMNS) / 2


def actual_header(frame, header):
    """The column in `frame` that answers to `header`, or None."""
    for name in RAW_ALIASES.get(header, [header]):
        if name in frame.columns:
            return name
    return None


def missing_raw_columns(frame):
    """Which raw columns are absent, named as they appear in the sheet."""
    missing = []
    for header in RAW_COLUMNS.values():
        if actual_header(frame, header) is None:
            missing.append(header)
    return missing


def raw_columns_found(frame):
    """How many of the raw columns this sheet has."""
    found = 0
    for header in RAW_COLUMNS.values():
        if actual_header(frame, header) is not None:
            found = found + 1
    return found


# ---------------------------------------------------------------------------
# The conversions the formulas used to do
# ---------------------------------------------------------------------------

NBSP = " "


def as_text(value):
    """
    A cell as clean text.

    Strips, and turns non-breaking spaces into ordinary ones. Pasted cells
    are full of them - 608 of 727 places of birth in the live register carry
    one - and they are invisible in Excel, print as a space, and break every
    exact comparison. Cleaning them is not worth telling anyone about.
    """
    if value is None or pd.isna(value):
        return ""
    return str(value).replace(NBSP, " ").strip()


def to_date(value):
    """
    A real date, or None when the cell cannot be read as one.

    Cells that Excel already stores as dates are used as they are. Cells
    typed as text are read day-first, and only with "/" or "-" between the
    parts, so 03/04/2025 is always the 3rd of April. The parser is never
    allowed to guess: a silent day-month swap would be invisible on a
    finished certificate and impossible to find later.

    Anything else - dots, written months, half-typed dates - comes back as
    None and the checks stop the run.
    """
    if value is None or pd.isna(value):
        return None
    if hasattr(value, "strftime"):
        return value

    text = as_text(value)
    if not text:
        return None

    for separator in config.DATE_SEPARATORS:
        if separator in text:
            parts = text.split(separator)
            if len(parts) != 3:
                return None
            try:
                day, month, year = (int(part) for part in parts)
            except ValueError:
                return None
            if year < 100:
                year = 2000 + year
            try:
                return datetime.date(year, month, day)
            except ValueError:
                return None

    return None


def as_date(value, pattern=None):
    """A date written the way certificates write them, or blank."""
    date = to_date(value)
    if date is None:
        return ""
    return date.strftime(pattern or config.DATE_FORMAT)


def full_sex(value):
    """
    Anything starting F or M, in any case, becomes FEMALE or MALE.

    The spreadsheet wrote IF(cell="F","FEMALE","MALE"), so a blank, a
    lowercase "f", or a typo all came out as MALE and printed. Here only a
    recognisable letter is accepted and everything else comes back blank,
    which the mandatory-field check then stops.
    """
    letter = as_text(value).upper()
    if letter.startswith("F"):
        return "FEMALE"
    if letter.startswith("M"):
        return "MALE"
    return ""


def serial_for(number, registration_date):
    """
    "123" and a 2026 registration date -> "123/2026".

    The formula had the year typed into it, so every certificate printed the
    year the formula was written rather than the year of the registration.
    Taking it from the registration date is what the column was always
    meant to say, and it stops being wrong every January.
    """
    number = as_text(number)
    if not number:
        return ""

    year = ""
    if registration_date is not None and not pd.isna(registration_date):
        if hasattr(registration_date, "year"):
            year = str(registration_date.year)

    if not year:
        return number
    return number + "/" + year


def joined(*values):
    """Join cells with a comma, the way the formulas concatenated them."""
    return ",".join(as_text(value) for value in values)


def joined_present(*values):
    """
    Join only the cells that have something in them.

    An address does not always use all three lines. The formula joined them
    regardless, so an applicant who wrote everything on the first line got
    "FULL ADDRESS,," printed on the certificate. Skipping the empty ones
    leaves the address reading the way it was written.
    """
    filled = [as_text(value) for value in values]
    return ",".join(part for part in filled if part)


# ---------------------------------------------------------------------------
# The transformation
# ---------------------------------------------------------------------------

# The raw columns carried through next to the finished ones. They feed the
# three columns built by joining cells together, and a join is never blank
# as long as one part is filled - so a missing address line or a missing
# date of birth hides inside the result and prints as a stray comma.
#
# Carrying them means the existing mandatory-field check sees the parts as
# well as the whole, reports the column by the name it has in the sheet,
# and does it after the finished rows have been filtered out.
CARRIED_COLUMNS = [
    "Date of birth",
    "Place of birth",
    "ADD 1",
    "ADD 2",
    "ADD 3",
    "MHA_File_No",
    "MHA_Date",
]


def transform(frame):
    """
    Build the certificate columns from a raw entry sheet.

    Returns a new table. Nothing is written back: the sheet keeps the raw
    columns exactly as they were typed, which is the point of moving the
    rules in here.
    """
    column = {}
    for key, header in RAW_COLUMNS.items():
        column[key] = frame[actual_header(frame, header)]

    prepared = pd.DataFrame(index=frame.index)

    prepared["Serial"] = [
        serial_for(number, date)
        for number, date in zip(column["serial"], column["registration_date"])
    ]
    prepared["File_Number"] = [as_text(v) for v in column["file_number"]]
    prepared["Name"] = [as_text(v).upper() for v in column["name"]]
    prepared["Sex"] = [full_sex(v) for v in column["sex"]]
    prepared["Name_of_the_Father"] = [as_text(v).upper() for v in column["father"]]
    prepared["Name_of_the_Mother"] = [as_text(v).upper() for v in column["mother"]]

    prepared["When_and_where_born"] = [
        joined(as_date(born), as_text(place).upper())
        for born, place in zip(column["birth_date"], column["place"])
    ]

    prepared["Description_and_residence_of_informant"] = [
        joined_present(one, two, three).upper()
        for one, two, three in zip(
            column["address_1"], column["address_2"], column["address_3"]
        )
    ]

    # Remember what could not be read, before the value is gone.
    #
    # to_date turns an unreadable cell into None, which from here on is
    # indistinguishable from a cell that was always empty - so the checks
    # reported "8.11.26" as a missing value and sent the officer to look at a
    # cell with something plainly in it. Recording the original text here is
    # what lets them say "date unreadable" and show it.
    unreadable = {}

    def dated(values, column_name):
        out = []
        for position, value in enumerate(values):
            converted = to_date(value)
            if converted is None and as_text(value):
                unreadable[(position, column_name)] = as_text(value)
            out.append(converted)
        return out

    prepared["Registration_date"] = dated(column["registration_date"],
                                          "Registration_date")

    prepared["MHA_File_And_date"] = [
        joined(as_text(number), as_date(date, "%d/%m/%y"))
        for number, date in zip(column["mha_file_number"], column["mha_date"])
    ]

    # Progress columns: keep what the sheet already records, and supply a
    # starting value for the rest. A sheet nobody has run yet has none of
    # them, so every record is pending, which is right.
    for name in STATE_COLUMNS:
        source = actual_header(frame, name)
        if source is not None:
            prepared[name] = frame[source].values
        elif name == "STATUS":
            prepared[name] = config.PENDING_STATUS
        elif name == "Signing_Authority_Name":
            # Left blank on purpose. The officer who signs a certificate is
            # the one thing nobody should be able to set by accident, so it
            # comes from a column in the sheet or the run stops.
            prepared[name] = None
        else:
            prepared[name] = None

    # The parts, so the checks can see them. Dates are carried as dates, not
    # as whatever was typed, so an unreadable one reads as blank here and
    # the mandatory-field check catches it instead of it reaching a
    # certificate as an empty space before a comma.
    # Everything that prints in capitals, in one place.
    for name in UPPERCASE_COLUMNS:
        if name in prepared.columns:
            prepared[name] = [
                as_text(value).upper() or None for value in prepared[name]
            ]

    # A template is built against the heading the sheet uses, so the signing
    # authority is also carried under that heading when it is not the
    # standard one. Otherwise the field is read but prints blank.
    authority = actual_header(frame, "Signing_Authority_Name")
    if authority is not None and authority not in prepared.columns:
        prepared[authority] = prepared["Signing_Authority_Name"]

    for header in CARRIED_COLUMNS:
        header = actual_header(frame, header) or header
        if header in ("Date of birth", "MHA_Date"):
            prepared[header] = dated(frame[header], header)
        else:
            prepared[header] = [as_text(v) for v in frame[header]]

    # Carried on the frame so the checks can tell an unreadable date from an
    # empty cell. attrs survives the copying and row filtering the pipeline
    # does, and the keys are positions in this frame, which the filtering
    # preserves as the index.
    prepared.attrs["unreadable_dates"] = unreadable
    return prepared


def whole_numbers(column):
    """
    True for a column of numbers with nothing after the point.

    pandas reads a number column that has one blank cell in it as decimals,
    so roll number 12 would print as "12.0".
    """
    if not pd.api.types.is_float_dtype(column):
        return False
    filled = column.dropna()
    return not filled.empty and bool((filled == filled.round()).all())


def prepare_plain(frame):
    """
    Hand back a case's sheet with its own columns, ready for the checks.

    Used for every case except the birth register (DOCUMATE_DATA_FORMAT=plain,
    the default). The header row becomes the field
    names in the template, so nothing is renamed or combined. Only three
    things are done to it:

      - text is stripped, and non-breaking spaces become ordinary ones,
        for the same reason as as_text()
      - whole numbers stay whole numbers instead of printing as 12.0
      - a STATUS column is supplied when the sheet has none, so every row
        starts pending; it is added to the sheet itself on the first write

    The ID column is checked here, before anything is generated. Without it
    the documents would be made and only the status write-back would fail.
    """
    id_name = config.id_column()
    if id_name not in frame.columns:
        raise ValueError(messages.NO_ID_COLUMN.format(
            column=id_name, columns=", ".join(str(c) for c in frame.columns)))

    prepared = frame.copy()

    for name in prepared.columns:
        column = prepared[name]
        if whole_numbers(column):
            prepared[name] = column.astype("Int64")
        elif column.dtype == object:
            prepared[name] = [
                as_text(value) if isinstance(value, str) else value
                for value in column
            ]

    if "STATUS" not in prepared.columns:
        prepared["STATUS"] = config.PENDING_STATUS

    return prepared


def prepare(frame):
    """
    Hand back a table in certificate shape, whatever shape came in.

    Only a case with DOCUMATE_DATA_FORMAT=birth_register is converted. Every
    other case keeps its own columns and goes through prepare_plain().

    For the birth register there are three outcomes, and the middle one is
    the important one:

    A sheet that is already prepared is returned untouched.

    A sheet that has most of the raw columns but not all of them is a raw
    sheet with something wrong - a renamed heading, usually - and it stops
    the run naming what is missing. Passing it through instead would fill
    the template with whatever columns happened to match and print
    certificates that are blank in half their fields, which is far worse
    than refusing.

    A sheet with none of the raw columns is something else entirely, so it
    is passed through and the ordinary checks describe it.
    """
    if not config.is_register():
        return prepare_plain(frame)

    if looks_prepared(frame):
        return frame

    missing = missing_raw_columns(frame)
    if not missing:
        return transform(frame)

    # "Nearly a raw sheet" means most of the columns are there. One or two
    # on their own prove nothing - "Serial" alone is as likely to belong to
    # a prepared sheet, or to a table that has nothing to do with us - so
    # only a clear majority is treated as a raw sheet with a fault in it.
    found = raw_columns_found(frame)
    if found < len(RAW_COLUMNS) / 2:
        return frame

    raise ValueError(messages.RAW_SHEET_INCOMPLETE.format(
        found=found,
        total=len(RAW_COLUMNS),
        missing=", ".join(missing),
    ))
