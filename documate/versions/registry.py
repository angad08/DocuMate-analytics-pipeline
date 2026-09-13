"""
The list of versions. THIS is the file you edit.

Every version of DocuMate is one entry below. It says which data source to
read from, which engine makes the Word file, and a few settings. There is
no code for each version - the parts are shared and the differences are
just these settings.

Want a new version? Add one entry. No new file, no copied code.

    key               short name you type on the command line
    label             how the version calls itself in messages
    source            "excel" or "postgres"
    engine            "docxtpl" (Python fills it) or "mailmerge" (Word fills it)
    template          which template file in templates\\
    output_prefix     start of the output filename
    check_records     run the v3 checks
                      (the database versions filter in SQL, so they don't)
    date_columns      which columns to turn into dd/mm/yyyy
                      (the database versions do that when reading, so none)
    batch_size        mailmerge only: records per run of Word.
                      None means one single run.
    max_workers       docxtpl only: how many records to fill at once
    timestamp_format  what gets stuck on the end of the filename
    poll              keep checking for new records by default
    poll_interval     how many seconds between checks
"""

import os

from documate.setup import config


# The two template files.
TEMPLATE_DOCXTPL = "DOCUMENT_TEMPLATE_FILE.docx"
TEMPLATE_MAILMERGE = "DOCUMENT_TEMPLATE_FILE_MM.docx"


class Version:
    """One version of DocuMate: which parts it uses, and its settings."""

    def __init__(self,
                 key,
                 label,
                 source,
                 engine,
                 template,
                 output_prefix,
                 check_records=True,
                 date_columns=None,
                 batch_size=None,
                 max_workers=10,
                 timestamp_format="%d%m%Y",
                 poll=False,
                 poll_interval=300,
                 notes=""):

        self.key = key
        self.label = label
        self.source = source
        self.engine = engine
        self.template = template
        self.output_prefix = output_prefix
        self.check_records = check_records
        self.date_columns = date_columns or []
        self.batch_size = batch_size
        self.max_workers = max_workers
        self.timestamp_format = timestamp_format
        self.poll = poll
        self.poll_interval = poll_interval
        self.notes = notes

    def template_path(self):
        return os.path.join(config.templates_folder(), self.template)

    def output_folder(self):
        return config.output_folder()


# ---------------------------------------------------------------------------
# The versions
# ---------------------------------------------------------------------------

VERSIONS = {}


# v3 - the one that went into production in 2025.
# First version with the checks, one merged file, and writing statuses back.
# Everything after this is the same thing with a different source or engine.
VERSIONS["v3"] = Version(
    key="v3",
    label="v3",
    source="excel",
    engine="docxtpl",
    template=TEMPLATE_DOCXTPL,
    output_prefix="DocuMatePLUS",
    check_records=True,
    date_columns=["Registration_date"],
    max_workers=10,
    notes="Excel + docxtpl. The production version.",
)


# X - same Excel front end as v3, but Word does the work instead of docxtpl.
# Batching is what stopped Word hanging on 1,440 records.
VERSIONS["x"] = Version(
    key="x",
    label="X",
    source="excel",
    engine="mailmerge",
    template=TEMPLATE_MAILMERGE,
    output_prefix="DocuMateX",
    check_records=True,
    date_columns=["Registration_date"],
    batch_size=250,
    notes="Excel + Word Mail Merge, 250 at a time. Fastest version.",
)


# Y - X with batching switched off. Slightly quicker on small runs,
# hangs on big ones, which is exactly why X exists.
VERSIONS["y"] = Version(
    key="y",
    label="Y",
    source="excel",
    engine="mailmerge",
    template=TEMPLATE_MAILMERGE,
    output_prefix="DocuMateY",
    check_records=True,
    date_columns=["Registration_date"],
    batch_size=None,
    notes="Excel + Word Mail Merge, all in one go. Small runs only.",
)


# Z - v3's engine, reading from PostgreSQL instead of Excel.
# The filtering moves into the SQL, so the checks are off and the dates are
# formatted while reading.
VERSIONS["z"] = Version(
    key="z",
    label="Z",
    source="postgres",
    engine="docxtpl",
    template=TEMPLATE_DOCXTPL,
    output_prefix="DocuMateZ",
    check_records=False,
    date_columns=[],
    max_workers=10,
    timestamp_format="%d%m%Y_%H%M%S",
    poll=True,
    poll_interval=30,
    notes="PostgreSQL + docxtpl, keeps checking for new records.",
)


# O - the last combination: Z's source with X's engine.
VERSIONS["o"] = Version(
    key="o",
    label="O",
    source="postgres",
    engine="mailmerge",
    template=TEMPLATE_MAILMERGE,
    output_prefix="DocuMateO",
    check_records=False,
    date_columns=[],
    batch_size=100,
    timestamp_format="%d%m%Y_%H%M%S",
    poll=True,
    poll_interval=30,
    notes="PostgreSQL + Word Mail Merge, 100 at a time.",
)


def get(key):
    """Find a version by name, listing the real ones if you mistype."""
    key = key.lower()

    if key not in VERSIONS:
        available = ", ".join(sorted(VERSIONS))
        raise SystemExit("Unknown version '" + key + "'. Available: " + available)

    return VERSIONS[key]
