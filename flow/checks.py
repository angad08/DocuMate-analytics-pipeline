"""
The checks you added in v3, plus the tidying that goes with them.

This is the part v3 introduced and every version after it inherited. It is
plain pandas - no Word, no database, no Windows - which is why it is the
part that can actually be tested.

When a check fails we raise CheckFailed with the message to show. The old
code signalled the same thing by setting self.data = None and returning,
which was easy to forget to look for.
"""

import pandas as pd

from setup import config
from setup import messages


class CheckFailed(Exception):
    """Raised when a check fails. Carries the message to show the user."""

    def __init__(self, message, extra_lines=None):
        Exception.__init__(self, message)
        self.message = message
        self.extra_lines = extra_lines or []


def is_blank(frame):
    """True wherever a cell is empty or only spaces."""
    empty = frame.isna()
    only_spaces = frame.apply(lambda column: column.astype(str).str.strip() == "")
    return empty | only_spaces


def drop_finished(data):
    """
    Keep only the rows that still need doing.

    Blank, "IN PROCESS", or anything else all count as pending. Only an
    exact match on PRINTED is dropped, ignoring case and spaces. That is
    why a half-typed status never causes a record to be skipped quietly.
    """
    status = data["STATUS"].fillna("").astype(str).str.strip().str.upper()
    return data[status != config.DONE_STATUS]


def check_records(data):
    """
    Run the five checks and hand back the rows worth processing.

        1. Is there any data at all?
        2. Is there a STATUS column?
        3. Is anything still pending?
        4. Are any must-have fields empty?
        5. Is the same File_Number in there twice?

    The first failure stops everything. The order matters: check 4 runs on
    the already-filtered rows, so an empty cell in a row that is finished
    and irrelevant can't block the run.
    """

    # --- 1. Any data at all? ---
    if data is None or data.empty:
        raise CheckFailed(messages.NO_DATA)

    # --- 2. Is there a STATUS column? ---
    if "STATUS" not in data.columns:
        raise CheckFailed(messages.NO_STATUS_COLUMN)

    # --- 3. Anything left to do? ---
    data = drop_finished(data)
    if data.empty:
        raise CheckFailed(messages.NO_PENDING)

    # --- 4. Are any must-have fields empty? ---
    # STATUS, Date_Printed and Date_Issued are allowed to be empty. Every
    # other column ends up in the certificate, so an empty one there means
    # a broken document.
    must_have = []
    for column in data.columns:
        if column not in config.OPTIONAL_COLUMNS:
            must_have.append(column)

    if must_have:
        bad_rows = data[is_blank(data[must_have]).any(axis=1)]

        if not bad_rows.empty:
            lines = [messages.MISSING_FIELDS_HEADER]

            for index, row in bad_rows.iterrows():
                empty_columns = []
                for column in must_have:
                    value = row[column]
                    if pd.isna(value) or str(value).strip() == "":
                        empty_columns.append(column)

                # +2 because row 1 is the header and pandas counts from 0,
                # so this prints the row number you see in Excel.
                lines.append(messages.MISSING_FIELDS_ROW.format(
                    row=index + 2,
                    fields=", ".join(empty_columns),
                ))

            raise CheckFailed(
                messages.MISSING_FIELDS.format(count=len(bad_rows)),
                lines,
            )

    # --- 5. Same applicant entered twice? ---
    key = config.DUPLICATE_KEY
    if key in data.columns:
        duplicates = data[data.duplicated(subset=[key], keep=False)]

        if not duplicates.empty:
            show_columns = []
            for column in [key, "Serial", "Name", "Date_Issued"]:
                if column in data.columns:
                    show_columns.append(column)

            raise CheckFailed(
                messages.DUPLICATES.format(count=len(duplicates), key=key),
                [duplicates[show_columns].to_string()],
            )

    return data


def format_dates(data, date_columns):
    """
    Turn date columns into dd/mm/yyyy so they look right in the document.

    A date pandas can't read becomes NaT, and the old code printed the word
    "NaT" straight into the certificate. We leave it empty instead.
    """
    data = data.copy()

    for column in date_columns:
        if column in data.columns:
            as_dates = pd.to_datetime(data[column], errors="coerce")
            data[column] = as_dates.dt.strftime("%d/%m/%Y").fillna("")

    return data


def sort_by_serial(data):
    """
    Sort by the number inside Serial, so the document prints in order.

    "123/2025" sorts as 123. Something with no number in it sorts as 0. If
    the whole column can't be read as numbers we skip the sort rather than
    fail the run, same as before.
    """
    column = config.SORT_COLUMN
    if column not in data.columns:
        return data

    try:
        numbers = data[column].astype(str).str.extract(r"(\d+)", expand=False)
        numbers = numbers.fillna("0").astype(int)

        data = data.assign(sort_key=numbers)
        data = data.sort_values("sort_key")
        return data.drop(columns="sort_key")

    except Exception:
        print(messages.SORT_SKIPPED)
        return data


def to_records(data, add_year=None):
    """
    Turn the table into the list of dictionaries the engines want.

    We use to_dict("records") rather than itertuples(). itertuples quietly
    renames any column that isn't a valid Python name to _1, _2 and so on,
    so renaming a column to have a space in it would blank that field in
    the certificate without any error.
    """
    from datetime import datetime

    if add_year is None:
        add_year = config.INJECT_YEAR

    records = data.to_dict("records")

    if add_year:
        year = datetime.now().year
        for record in records:
            record.setdefault("year", year)

    return records
