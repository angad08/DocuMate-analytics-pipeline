"""
The list of DocuMate versions, and the settings for each one.

This is the file to edit when you want to change how a version behaves or
add a new one. Each version is one Version(...) entry below. There is no
separate code per version: every version uses the same shared parts, and
these settings choose which parts and how they are configured.

To add a version, copy an existing entry, give it a new key, and change the
settings. It is then available as  python main.py <key>.

Settings
--------
    key               name you type on the command line, e.g. "x"
    label             name shown in messages
    source            "excel" or "database"
                      ("database" uses Azure SQL or PostgreSQL, whichever
                       DOCUMATE_DB_BACKEND in config.env selects)
    engine            "docxtpl" (Python fills the template) or
                      "mailmerge" (Word fills it)
    template          None: the template chosen when the run starts. A
                      file name here ties the version to that one template.
    output_prefix     start of the output file name
    check_records     run the record checks in flow/checks.py
                      (off for database versions, which filter in SQL)
    date_columns      None: worked out from the sheet (Excel versions).
                      The database versions pass [], they format on read.

Nothing here says what is being generated - that is whichever Excel file
and Word template are chosen at the start - so every version works for any
kind of document.
    batch_size        mailmerge only: records per Word merge;
                      None merges everything at once
    max_workers       docxtpl only: records filled at the same time
    timestamp_format  date/time added to the end of the output file name
    poll              keep checking for new records by default
    poll_interval     seconds between checks when polling
    to_pdf            also save a PDF copy of the merged Word file, with
                      the same name in the same folder. Needs Word.
"""

from setup import config


class Version:
    """One DocuMate version: which source and engine it uses, and its settings."""

    def __init__(self,
                 key,
                 label,
                 source,
                 engine,
                 template=None,
                 output_prefix="DocuMate",
                 check_records=True,
                 date_columns=None,
                 batch_size=None,
                 max_workers=10,
                 timestamp_format="%d%m%Y",
                 poll=False,
                 poll_interval=300,
                 to_pdf=False,
                 notes=""):

        self.key = key
        self.label = label
        self.source = source
        self.engine = engine
        self.template = template
        self.output_prefix = output_prefix
        self.check_records = check_records
        self.own_date_columns = date_columns
        self.batch_size = batch_size
        self.max_workers = max_workers
        self.timestamp_format = timestamp_format
        self.poll = poll
        self.poll_interval = poll_interval
        self.to_pdf = to_pdf
        self.notes = notes

    @property
    def date_columns(self):
        """
        Columns to format as dd/mm/yyyy: the version's own list when it has
        one, otherwise the sheet's (see config.date_columns).
        """
        if self.own_date_columns is not None:
            return self.own_date_columns
        return config.date_columns()

    def template_path(self):
        """
        Full path to this version's template: the case's template for this
        engine, or the version's own file when it names one.

        Raises ValueError naming the setting when the case has not set it.
        """
        if self.template:
            return config.in_folder(config.templates_folder(), self.template)
        return config.template_path(self.engine)

    def output_folder(self):
        """Folder the output files are saved in."""
        return config.output_folder()


# ---------------------------------------------------------------------------
# The versions
# ---------------------------------------------------------------------------

VERSIONS = {}


# v3: Excel + docxtpl. The production version.
# Checks the records, merges them into one file, and marks them PRINTED.
VERSIONS["v3"] = Version(
    key="v3",
    label="v3",
    source="excel",
    engine="docxtpl",
    output_prefix="DocuMatePLUS",
    check_records=True,
    max_workers=10,
    to_pdf=True,
    notes="Excel + docxtpl. The production version.",
)


# X: Excel + Word Mail Merge, 250 records per merge.
# Batching keeps Word responsive on large runs.
VERSIONS["x"] = Version(
    key="x",
    label="X",
    source="excel",
    engine="mailmerge",
    output_prefix="DocuMateX",
    check_records=True,
    batch_size=250,
    to_pdf=True,
    notes="Excel + Word Mail Merge, 250 at a time. Fastest version.",
)


# Y: the same as X but merges everything in one go (no batching).
# Fine for small runs; Word can stop responding on large ones.
VERSIONS["y"] = Version(
    key="y",
    label="Y",
    source="excel",
    engine="mailmerge",
    output_prefix="DocuMateY",
    check_records=True,
    batch_size=None,
    to_pdf=True,
    notes="Excel + Word Mail Merge, all in one go. Small runs only.",
)


# Z: database + docxtpl.
# The SQL only returns pending records and the dates are formatted on read,
# so checks are off and date_columns is empty. Polls every 30 seconds.
VERSIONS["z"] = Version(
    key="z",
    label="Z",
    source="database",
    engine="docxtpl",
    output_prefix="DocuMateZ",
    check_records=False,
    date_columns=[],
    max_workers=10,
    timestamp_format="%d%m%Y_%H%M%S",
    poll=True,
    poll_interval=30,
    to_pdf=True,
    notes="Database + docxtpl, keeps checking for new records.",
)


# O: database + Word Mail Merge, 100 records per merge. Polls every 30 seconds.
VERSIONS["o"] = Version(
    key="o",
    label="O",
    source="database",
    engine="mailmerge",
    output_prefix="DocuMateO",
    check_records=False,
    date_columns=[],
    batch_size=100,
    timestamp_format="%d%m%Y_%H%M%S",
    poll=True,
    poll_interval=30,
    to_pdf=True,
    notes="Database + Word Mail Merge, 100 at a time.",
)


def get(key):
    """
    Return the version with this key (case-insensitive).

    Stops with a message listing the valid keys if it does not exist.
    """
    key = key.lower()

    if key not in VERSIONS:
        available = ", ".join(sorted(VERSIONS))
        raise SystemExit("Unknown version '" + key + "'. Available: " + available)

    return VERSIONS[key]
