"""
The checks you added in v3, plus the tidying that goes with them.

This is the part v3 introduced and every version after it inherited. It is
plain pandas - no Word, no database, no Windows - which is why it is the
part that can actually be tested.

When a check fails we raise CheckFailed with the message to show. The old
code signalled the same thing by setting self.data = None and returning,
which was easy to forget to look for.
"""

import datetime
import re

import pandas as pd

from setup import config
from setup import messages
from sources import excel_format


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
    exact match on COMPLETED or PRINTED is dropped, ignoring case and
    spaces. That is why a half-typed status never causes a record to be
    skipped quietly.

    COMPLETED is dropped as well as PRINTED because its document has
    already been generated. Treating it as pending would build the same
    certificate again on the next run.
    """
    status = data["STATUS"].fillna("").astype(str).str.strip().str.upper()
    return data[~status.isin(config.FINISHED_STATUSES)]


def prefix_failures(data, column, prefix):
    """
    Rows whose value in `column` doesn't start with `prefix`.

    An empty prefix turns the check off. Real MHA file numbers carry no
    prefix at all - they are plain twelve-digit numbers - so that one is
    off by default and only the file number is checked.
    """
    if not prefix or column not in data.columns:
        return []

    failures = []
    for index, value in data[column].items():
        text = "" if pd.isna(value) else str(value).strip()
        if text and not text.startswith(prefix):
            failures.append((index + 2, text))
    return failures


def serial_numbers_missing(data):
    """Row numbers, as Excel shows them, where S No. is blank."""
    if "Serial" not in data.columns:
        return []

    missing = []
    for index, value in data["Serial"].items():
        if pd.isna(value) or str(value).strip() == "":
            missing.append(index + 2)
    return missing


def renumber_serials(data, year_column="Registration_date"):
    """
    Replace every Serial with its position in the sheet: 1/year .. n/year.

    Counts the rows and numbers them straight through, so a column with
    gaps in it comes out continuous. Row order is never changed - only the
    number written against each row - and nothing is saved back to the
    workbook, so the sheet still needs filling in properly afterwards.
    """
    data = data.copy()
    serials = []

    for position, (_, row) in enumerate(data.iterrows(), start=1):
        year = ""
        value = row.get(year_column)
        if value is not None and not pd.isna(value) and hasattr(value, "year"):
            year = str(value.year)
        serials.append(str(position) + "/" + year if year else str(position))

    data["Serial"] = serials
    return data


def reads_as_date(value):
    """
    True when `value` is a date the certificate can actually print.

    Delegates to excel_format.to_date rather than parsing again, so the check
    and the converter can never disagree. A second parser here would eventually
    accept something to_date rejects, the row would pass validation, and the
    certificate would print an empty date — the failure this check exists to
    stop.

    Which separators count is config.DATE_SEPARATORS, currently "/" and "-".
    Dots are excluded on purpose, so 8.11.26 is reported as a format problem
    rather than silently read. Add "." to that list to accept it.

    A time with no date is rejected separately: Excel hands back datetime.time
    for a cell holding only 00:00, and to_date lets it through because a time
    has .strftime — but it carries no day, so it cannot be a date of birth.
    """
    if isinstance(value, datetime.time) and not isinstance(value, datetime.datetime):
        return False
    return excel_format.to_date(value) is not None


def field_problems(data, must_have, date_columns=()):
    """
    Every cell that would break a certificate, as (row, column, issue, value).

    Two different problems, reported differently because the fix differs:

      Missing       - the cell is empty; someone has to supply the value.
      Date unreadable - the cell has something in it, but not a date this can
                      print. Calling that "missing" sends the officer looking
                      at a cell that is plainly not empty, which is what used
                      to happen with 8.11.26.

    `row` is the number shown in Excel: +2 because row 1 is the header and
    pandas counts from 0.
    """
    problems = []
    dates = {c for c in date_columns if c in data.columns}

    # What the raw-sheet conversion could not read. Those cells are already
    # None by now, so without this they would be reported as missing — which
    # is exactly the complaint: 8.11.26 is not an empty cell.
    unreadable = data.attrs.get("unreadable_dates", {})

    for index, row in data.iterrows():
        for column in must_have:
            value = row[column]
            blank = pd.isna(value) or str(value).strip() == ""
            original = unreadable.get((index, column))
            if blank and original:
                problems.append((index + 2, column, "Date unreadable", original))
            elif blank:
                problems.append((index + 2, column, "Missing", ""))
            elif column in dates and not reads_as_date(value):
                problems.append((index + 2, column, "Date unreadable",
                                 str(value).strip()))
    return problems


def case_rule_problems(data):
    """
    Every cell that breaks the case file's rules, as (row, column, issue, value).

    The rules are the case file's MUST_FILL, NO_REPEATS, ALLOWED,
    STARTS_WITH and DATES (see setup/cases.py). With no case file, or one
    that sets none of them, nothing is checked and the list is empty.

    Raises CheckFailed straight away if a rule names a column the sheet
    does not have: a check on a missing column would find nothing wrong and
    let every row through.

    Values are compared as the sheet shows them: spaces trimmed, and
    ALLOWED ignoring capitals, so "f" is as good as "F".
    """
    from setup import cases

    named = cases.named_columns()
    if not named:
        return []

    absent = [column for column in named if column not in data.columns]
    if absent:
        raise CheckFailed(messages.CASE_COLUMNS_MISSING.format(
            count=len(absent),
            columns=", ".join(absent),
            available=", ".join(str(c) for c in data.columns
                                if c not in config.STATE_COLUMNS_OPTIONAL),
        ))

    # Rows are marked done by matching their ID, so the ID has to reach the
    # write-back exactly as the sheet holds it. Printing it differently
    # would leave every row pending and make the same documents next run.
    for id_name in (config.REGISTER_ID_COLUMN if config.is_register() else "",
                    config.setting("DOCUMATE_ID_COLUMN", ""),
                    config.setting("DOCUMATE_DB_ID_COLUMN", "")):
        if id_name and (id_name in cases.capitals() or id_name in cases.replace()):
            raise CheckFailed(messages.CASE_CHANGES_ID.format(
                column=id_name, column_name=re.sub(r"\W+", "_", id_name).strip("_")))

    must_fill = cases.must_fill()
    dates = cases.dates()
    allowed = cases.allowed()
    starts_with = cases.starts_with()

    problems = []
    for index, row in data.iterrows():
        excel_row = index + 2
        for column in named:
            text = cell_text(row[column])

            if not text:
                if column in must_fill:
                    problems.append((excel_row, column, messages.ISSUE_MUST_FILL, ""))
                continue

            if column in dates and not reads_as_date(row[column]):
                problems.append((excel_row, column, messages.ISSUE_NOT_A_DATE, text))

            choices = allowed.get(column)
            if choices and text.upper() not in {choice.upper() for choice in choices}:
                problems.append((excel_row, column, messages.ISSUE_NOT_ALLOWED.format(
                    values=", ".join(choices)), text))

            prefix = starts_with.get(column)
            if prefix and not text.startswith(prefix):
                problems.append((excel_row, column, messages.ISSUE_WRONG_START.format(
                    prefix=prefix), text))

    # Repeats are found across rows, so after the row-by-row pass. Every
    # copy is listed, not just the second, because either could be the
    # mistake.
    for column in cases.no_repeats():
        texts = data[column].apply(cell_text)
        repeated = texts[(texts != "") & texts.duplicated(keep=False)]
        for index, text in repeated.items():
            problems.append((index + 2, column, messages.ISSUE_REPEATED, text))

    return problems


def apply_case_formatting(data):
    """
    How the case file says things should print: REPLACE, then JOIN, then
    CAPITALS (see setup/cases.py). Returns a new table; the sheet itself
    is never written to.

    The order is the one that reads naturally in a case file:

      REPLACE first, so F becomes FEMALE before anything is put together.
      JOIN next, leaving blank parts out - an address with no second line
          prints "10 ROAD, PERTH", not "10 ROAD, , PERTH". Dates are already
          dd/mm/yyyy by now, so a joined date reads like any other.
      CAPITALS last, so it can name a JOIN as well as a column.

    With no case file nothing changes.
    """
    from setup import cases

    swaps = cases.replace()
    joins = cases.joins()
    upper = cases.capitals()
    if not (swaps or joins or upper):
        return data

    data = data.copy()

    for column, pairs in swaps.items():
        data[column] = [pairs.get(cell_text(value).upper(), value)
                        for value in data[column]]

    separator = cases.join_separator()
    for name, parts in joins:
        data[name] = [separator.join(text for text in (cell_text(row[part]) for part in parts)
                                     if text)
                      for _index, row in data.iterrows()]

    for column in upper:
        # A blank stays blank rather than becoming the text "NAN".
        data[column] = [cell_text(value).upper() if cell_text(value) else value
                        for value in data[column]]

    return data


def cell_text(value):
    """A cell as the sheet shows it: trimmed, and 12.0 written as 12."""
    if isinstance(value, float) and not pd.isna(value) and value.is_integer():
        return str(int(value))
    return excel_format.as_text(value)


def problem_table(problems):
    """
    The problem list as aligned columns: Row, Field, Issue, Value.

    A row per problem rather than per record, so a record with two different
    problems says so on two lines instead of merging them into one phrase that
    hides which field had which. Sorted by row, then field, so it reads in the
    order the sheet does.

    Returns a list of lines, which is what CheckFailed carries.
    """
    header = ("Row", "Field", "Issue", "Value")
    body = sorted(((str(r), c, i, v) for r, c, i, v in problems),
                  key=lambda t: (int(t[0]), t[1]))

    widths = [max(len(row[n]) for row in (header, *body)) for n in range(4)]
    rule = "  ".join("-" * w for w in widths)

    def line(cells):
        return "  ".join(text.ljust(w) for text, w in zip(cells, widths)).rstrip()

    lines = [messages.MISSING_FIELDS_HEADER, line(header), rule]
    lines += [line(row) for row in body]
    return lines


def check_records(data, serial_fixer=None, date_columns=(), required=None,
                  blanks_ok=None):
    """
    Run the five checks and hand back the rows worth processing.

    required    the columns the template prints, for any sheet but the
                birth register; None checks every column (the register)
    blanks_ok   asked, with the number of rows, whether blank cells in
                `required` columns may print as blanks; None refuses them

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

    # --- 3b. Blank S No. on a raw entry sheet ---
    # Only a raw sheet has this column, so a prepared one skips it. It runs
    # after the pending filter, so a finished row with a gap in it cannot
    # block a run. What to do about the gaps is the caller's decision: the
    # pipeline asks when someone is watching and refuses when it is not.
    # Another case study's ID column is checked like any other field in 4:
    # numbering it 1/year would invent IDs that mean nothing in that sheet.
    missing_serials = serial_numbers_missing(data) if config.is_register() else []
    if missing_serials:
        if serial_fixer is None:
            raise CheckFailed(
                messages.SERIAL_REFUSED.format(count=len(missing_serials)))
        data = serial_fixer(data, missing_serials)

    # --- 3c. Is there a signing authority at all? ---
    # Checked on its own because the fix is a setting rather than a row
    # edit, and because an officer's name on an official document is not
    # something to supply from a default.
    if config.is_register() and "Signing_Authority_Name" in data.columns:
        filled = data["Signing_Authority_Name"].apply(
            lambda v: not (pd.isna(v) or str(v).strip() == ""))
        if not filled.any():
            raise CheckFailed(messages.NO_SIGNING_AUTHORITY)

    # --- 3d. The case's own rules ---
    # Only when a case file sets some. Before step 4 on purpose: a MUST_FILL
    # column is never one to be asked about, so its blanks are reported here
    # with everything else this case forbids, all in one list.
    case_problems = case_rule_problems(data)
    if case_problems:
        rows_affected = len({row for row, _c, _i, _v in case_problems})
        raise CheckFailed(messages.CASE_RULES_FAILED.format(count=rows_affected),
                          problem_table(case_problems))

    # --- 4. Are any must-have fields empty? ---
    # The birth register: every column except the optional ones, because
    # its columns are joined into certificate fields and a gap inside a join
    # would hide. Any other sheet: only the columns its template prints
    # (`required`), and the ID column, which every row is marked done by.
    if required is None:
        optional = config.optional_columns()
        must_have = [column for column in data.columns if column not in optional]
    else:
        must_have = list(required)
        id_name = config.id_column()
        if id_name in data.columns and id_name not in must_have:
            must_have.append(id_name)

    if must_have:
        problems = field_problems(data, must_have, date_columns)

        # A blank cell in a printed field is sometimes right - a Remarks
        # column - so with someone watching, `blanks_ok` may let it through.
        # Never for the ID, and never for a date that could not be read:
        # those are mistakes, not choices.
        if problems and required is not None and blanks_ok is not None:
            id_name = config.id_column()
            only_blanks = all(issue == "Missing" and column != id_name
                              for _r, column, issue, _v in problems)
            if only_blanks:
                rows_affected = len({row for row, _c, _i, _v in problems})
                for line in problem_table(problems):
                    print(line)
                if blanks_ok(rows_affected):
                    problems = []
                else:
                    # The rows are already on screen; listing them twice
                    # only pushes the question that was answered out of view.
                    raise CheckFailed(messages.MISSING_FIELDS.format(count=rows_affected))

        if problems:
            rows_affected = len({row for row, _c, _i, _v in problems})
            raise CheckFailed(
                messages.MISSING_FIELDS.format(count=rows_affected),
                problem_table(problems),
            )

    # --- 4b. Do the numbers look like ours? ---
    # Only runs on a raw sheet: a prepared one has no "File No" column, and
    # a file number issued elsewhere is either a typo or a record that does
    # not belong in this register. An empty prefix in config turns it off.
    for column, prefix, label in [
        ("File_Number", config.file_number_prefix(), "File No"),
        ("MHA_File_No", config.mha_file_prefix(), "MHA_File_No"),
    ]:
        failures = prefix_failures(data, column, prefix)
        if failures:
            lines = [messages.PREFIX_ROW.format(row=row, value=value)
                     for row, value in failures[:20]]
            raise CheckFailed(
                messages.PREFIX_MISMATCH.format(
                    count=len(failures), column=label, prefix=prefix),
                lines,
            )

    # --- 5. Same applicant entered twice? ---
    key = config.duplicate_column()
    if key and key in data.columns:
        duplicates = data[data.duplicated(subset=[key], keep=False)]

        if not duplicates.empty:
            show_columns = []
            for column in [key, config.id_column(), "Name", "Date_Issued"]:
                if column in data.columns and column not in show_columns:
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
            # Two passes, because one setting cannot read both shapes.
            #
            # Left to itself pandas falls back to dateutil, which reads
            # 05/06/2026 month-first - the 6th of May instead of the 5th of
            # June - and throws away 25/06/2026 entirely, there being no month
            # 25. Forcing dayfirst=True instead breaks the other shape:
            # 2026-01-05 then reads as the 1st of May.
            #
            # So ISO is parsed as ISO, and only what is left over is read
            # day-first, which is the house rule everywhere else (to_date,
            # config.DATE_SEPARATORS). A real date object matches the first
            # pass and is untouched by either.
            #
            # It matters because a swapped date looks completely ordinary on a
            # finished certificate - there is nothing to notice later.
            column_values = data[column]
            as_dates = pd.to_datetime(column_values, format="ISO8601",
                                      errors="coerce")
            if as_dates.isna().any():
                day_first = pd.to_datetime(column_values.where(as_dates.isna()),
                                           errors="coerce", dayfirst=True)
                as_dates = as_dates.fillna(day_first)
            data[column] = as_dates.dt.strftime("%d/%m/%Y").fillna("")

    return data


def sort_by_serial(data):
    """
    Sort by the number inside Serial, so the document prints in order.

    "123/2025" sorts as 123. Something with no number in it sorts as 0. If
    the whole column can't be read as numbers we skip the sort rather than
    fail the run, same as before.
    """
    column = config.sort_column()
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

    # An empty optional cell reaches here as NaN, NA or None, and docxtpl
    # prints each of those as the word itself - "nan" on the document. A
    # blank cell should print as nothing.
    for record in records:
        for name, value in record.items():
            if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)):
                record[name] = ""

    if add_year:
        year = datetime.now().year
        for record in records:
            record.setdefault("year", year)

    return records
