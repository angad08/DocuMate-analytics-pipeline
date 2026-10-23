"""
Case files: one small settings file per kind of document.

    python main.py v3 foreign_pcc

reads cases/foreign_pcc.env before anything else. It is written exactly
like config.env - one setting per line - and holds two kinds of line:

    The files, under the names config.env already uses:

        DOCUMATE_EXCEL_PATH=PCC_2026.xlsx
        DOCUMATE_TEMPLATE=PCC_FORMAT.docx
        DOCUMATE_ID_COLUMN=Passport No

    The case's own checks, named in plain words:

        MUST_FILL=Passport No, Applicant Name
        NO_REPEATS=Passport No, File No
        ALLOWED=Sex: F, M ; Grade: A, B, C
        STARTS_WITH=File No: IND/MEL/
        DATES=Issue Date, Date of birth

    And how things print, applied once the checks have passed - to the
    documents only, never to the sheet:

        REPLACE=Sex: F=FEMALE, M=MALE
        JOIN=Address: ADD 1 + ADD 2 + ADD 3
        JOIN_SEPARATOR=", "
        CAPITALS=Applicant Name, Address

Every line is optional. A case file's settings win over config.env's, so
config.env can stay the machine's settings - the database, the switches -
while each case brings its own files and rules. With no case named,
DocuMate runs exactly as it always has.

The answers go into the environment, like everything the picker settles,
so the rest of DocuMate reads them the same way it reads config.env.
"""

import os
import re

from setup import config
from setup import messages


# The config.env settings a case file may set: the ones that describe a
# kind of document rather than the machine it runs on.
FILE_SETTINGS = [
    "DOCUMATE_EXCEL_PATH",
    "DOCUMATE_SHEET_NAME",
    "DOCUMATE_TEMPLATE",
    "DOCUMATE_TEMPLATE_MM",
    "DOCUMATE_ID_COLUMN",
    "DOCUMATE_OUTPUT_NAME",
    "DOCUMATE_DB_TABLE",
    "DOCUMATE_DB_ID_COLUMN",
]

# The case's own checks, and the environment name each is kept under. The
# CASE_ in the name keeps them from being mistaken for config.env settings.
RULES = {
    "MUST_FILL": "DOCUMATE_CASE_MUST_FILL",
    "NO_REPEATS": "DOCUMATE_CASE_NO_REPEATS",
    "ALLOWED": "DOCUMATE_CASE_ALLOWED",
    "STARTS_WITH": "DOCUMATE_CASE_STARTS_WITH",
    "DATES": "DOCUMATE_CASE_DATES",
    "REPLACE": "DOCUMATE_CASE_REPLACE",
    "JOIN": "DOCUMATE_CASE_JOIN",
    "JOIN_SEPARATOR": "DOCUMATE_CASE_JOIN_SEPARATOR",
    "CAPITALS": "DOCUMATE_CASE_CAPITALS",
}

# What goes between the parts of a JOIN when JOIN_SEPARATOR is not set.
DEFAULT_JOIN_SEPARATOR = ", "

# A JOIN's name is used as a {{ placeholder }}, so it has to be one.
PLACEHOLDER_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def cases_folder():
    """cases/ in the project folder, beside config.env."""
    return os.path.join(config.project_root(), "cases")


def available():
    """The names of the case files there are, for --list and messages."""
    folder = cases_folder()
    if not os.path.isdir(folder):
        return []
    return sorted(os.path.splitext(name)[0] for name in os.listdir(folder)
                  if name.lower().endswith(".env"))


def case_path(name):
    """cases/<name>.env, or a path to a .env file given as it is."""
    if name.lower().endswith(".env") and os.path.isfile(name):
        return name
    return os.path.join(cases_folder(), name + ".env")


def load(name):
    """
    Read a case file into the environment.

    Raises ValueError, with a message to show, when there is no such case
    or a line in it is not a setting DocuMate knows - a misspelt rule would
    otherwise be ignored and the check it names would quietly never run.
    """
    path = case_path(name)
    if not os.path.isfile(path):
        raise ValueError(messages.CASE_MISSING.format(
            name=name, folder=cases_folder(),
            cases=", ".join(available()) or "none yet"))

    try:
        from dotenv import dotenv_values
    except ImportError:
        raise ValueError(messages.CASE_NEEDS_DOTENV)

    # utf-8-sig for the same reason as config.env: Notepad can save an
    # invisible mark at the start that would hide the first line.
    values = dotenv_values(path, encoding="utf-8-sig")

    unknown = [key for key in values
               if key not in FILE_SETTINGS and key not in RULES]
    if unknown:
        raise ValueError(messages.CASE_UNKNOWN_SETTING.format(
            path=path, unknown=", ".join(unknown),
            known=", ".join(FILE_SETTINGS + list(RULES))))

    for key, value in values.items():
        # A blank line counts as not set, as it does in config.env.
        if not value or not value.strip():
            continue
        # The separator is kept exactly as written: its spaces are the
        # point of it, and quotes ( JOIN_SEPARATOR=", " ) keep them.
        if key != "JOIN_SEPARATOR":
            value = value.strip()
        os.environ[RULES.get(key, key)] = value

    os.environ["DOCUMATE_CASE"] = name

    # Read every rule once now, so a badly written one is reported when the
    # case is opened - naming the line - rather than part-way through a run.
    named_columns()
    return path


# ---------------------------------------------------------------------------
# Reading the rules
# ---------------------------------------------------------------------------

def column_list(text):
    """"Passport No, File No" -> ["Passport No", "File No"]."""
    return [part.strip() for part in text.split(",") if part.strip()]


def column_values(text, setting):
    """
    "Sex: F, M ; Grade: A, B" -> {"Sex": ["F", "M"], "Grade": ["A", "B"]}.

    One entry per column, separated by ";". Raises ValueError naming the
    setting when an entry has no ":".
    """
    found = {}
    for entry in text.split(";"):
        if not entry.strip():
            continue
        if ":" not in entry:
            raise ValueError(messages.CASE_BAD_RULE.format(setting=setting, entry=entry.strip()))
        column, values = entry.split(":", 1)
        found[column.strip()] = column_list(values)
    return found


def must_fill():
    return column_list(config.setting(RULES["MUST_FILL"], ""))


def no_repeats():
    return column_list(config.setting(RULES["NO_REPEATS"], ""))


def allowed():
    return column_values(config.setting(RULES["ALLOWED"], ""), "ALLOWED")


def starts_with():
    """{column: prefix}. Each column has one prefix, so only the first is kept."""
    found = column_values(config.setting(RULES["STARTS_WITH"], ""), "STARTS_WITH")
    return {column: prefixes[0] for column, prefixes in found.items() if prefixes}


def dates():
    return column_list(config.setting(RULES["DATES"], ""))


def replace():
    """
    "Sex: F=FEMALE, M=MALE" -> {"Sex": {"F": "FEMALE", "M": "MALE"}}.

    The typed side is kept in capitals so "f" matches F; values not listed
    are printed as typed.
    """
    found = {}
    for column, pairs in column_values(config.setting(RULES["REPLACE"], ""), "REPLACE").items():
        swaps = {}
        for pair in pairs:
            if "=" not in pair:
                raise ValueError(messages.CASE_BAD_REPLACE.format(entry=pair))
            typed, shown = pair.split("=", 1)
            swaps[typed.strip().upper()] = shown.strip()
        found[column] = swaps
    return found


def joins():
    """
    "Address: ADD 1 + ADD 2 ; Born: Date of birth + Place"
        -> [("Address", ["ADD 1", "ADD 2"]), ("Born", ["Date of birth", "Place"])]

    In the order written, so a JOIN can use one made before it.
    """
    found = []
    for name, parts in column_values(config.setting(RULES["JOIN"], ""), "JOIN").items():
        if not PLACEHOLDER_NAME.match(name):
            raise ValueError(messages.CASE_BAD_JOIN_NAME.format(
                name=name, suggestion=re.sub(r"\W+", "_", name).strip("_") or "Address"))
        # column_values splits on commas; a JOIN's parts are joined by +.
        columns = [part.strip() for part in ",".join(parts).split("+") if part.strip()]
        if not columns:
            raise ValueError(messages.CASE_BAD_JOIN.format(name=name))
        found.append((name, columns))
    return found


def join_names():
    return [name for name, _parts in joins()]


def join_separator():
    """JOIN_SEPARATOR exactly as written, spaces and all; ", " when not set."""
    config.load_environment()
    return os.environ.get(RULES["JOIN_SEPARATOR"]) or DEFAULT_JOIN_SEPARATOR


def capitals():
    return column_list(config.setting(RULES["CAPITALS"], ""))


def named_columns():
    """
    Every column of the sheet any rule names, so one missing from the
    sheet can be caught. A JOIN's own name is not a column, so it is left
    out wherever it is used.
    """
    made = set(join_names())
    named = must_fill() + no_repeats() + dates()
    named += list(allowed()) + list(starts_with()) + list(replace())
    named += [part for _name, parts in joins() for part in parts]
    named += capitals()
    return [column for column in dict.fromkeys(named) if column not in made]
