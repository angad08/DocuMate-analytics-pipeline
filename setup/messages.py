"""
Every message DocuMate shows the user, in one place.

Each old script had slightly different wording for the same thing -
"STATUSes" in one, "statuses" in another, "Powered by Word Mail Merge +
Python" in X and "+ PostgreSQL" in O. None of that told the user anything,
so it is one set now.

Want to change what DocuMate says? Change it here. Nowhere else has text.
"""

SIGNATURE = "Powered by ENN - Fuelled by curiosity, refined by data-driven clarity."

# --- Starting up -----------------------------------------------------------
BANNER = "Deploying DocuMate {version}...\n"
READY = "DocuMate : Alright ENN, Ready to engage - initializing populate launch\n"

# --- The main steps --------------------------------------------------------
READING = "DocuMate : Reading {source}..."
SEARCHING = "DocuMate : Searching for new records..."
FOUND = "DocuMate : Found {count} new records to populate.\n"
ENGINE_START = "DocuMate : {engine} starting...\n"
GENERATING = "DocuMate : Generating and merging {count} documents...\n"
RECORD_PROGRESS = "DocuMate : Processed record {done}/{total}"
CREATING_FOLDER = "\nDocuMate : Output folder '{folder}' does not exist. Creating it..."
SAVED = "\nDocuMate : Generated and merged {count} documents (in Serial order) and saved in:\n{path}"
MERGE_DONE = "\nDocuMate : Generated and merged {count} documents"
MERGE_TIME = "DocuMate : {engine} finished in {seconds:.2f} seconds.\n"
NOTHING_TO_MERGE = "DocuMate : No records found for merging."

# --- PDF copy --------------------------------------------------------------
PDF_START = "\nDocuMate : Exporting a PDF copy of the merged document..."
PDF_SAVED = "DocuMate : PDF saved in {seconds:.2f} seconds:\n{path}"
PDF_FAILED = "DocuMate : PDF export failed, the Word file is still saved - {error}"

# --- Mail Merge only -------------------------------------------------------
MM_SOURCE_BUILT = "DocuMate : Built CSV merge source - {count} records, 1 to {count}"
MM_STARTING = "DocuMate : Starting Word Mail Merge..."
MM_CONNECTING = "DocuMate : Connecting CSV data source, please wait..."
MM_CONNECTED = "DocuMate : Data source connected successfully."
MM_RANGE = "DocuMate : Merge range in Word: From {first} To {last}"
MM_BATCH = "DocuMate : Batch {number} of {total} - records {first} to {last}"
MM_ERROR = "\nDocuMate : Mail Merge error - {error}"
CLEANUP_FAILED = "DocuMate : Could not delete temp file: {path}"

# --- Marking records as printed --------------------------------------------
ASK_UPDATE = "DocuMate : Mark {count} as printed?{failed} (yes/no): "
EXCEL_NO_SERIAL_COLUMN = "DocuMate : Serial column not found - cannot update Excel."
UPDATING = "\nDocuMate : Updating {target} statuses...\n"
UPDATE_SUFFIX = " and updated {target} in {seconds:.2f} seconds"
EXCEL_UPDATED = "DocuMate : Sheet '{sheet}' updated successfully on {date} at {clock}."
EXCEL_NO_STATUS_COLUMN = "DocuMate : STATUS column not found - cannot update Excel."
EXCEL_LOCKED = "DocuMate : Cannot update Excel while the file is open. Close it and retry."
EXCEL_UPDATE_ERROR = "DocuMate : Error updating Excel - {error}"
EXCEL_COUNT_WARNING = "DocuMate : Expected to mark {expected} rows but marked {marked} - check the Serial values."
DB_UPDATED = "DocuMate : Database updated successfully ({count} records)."
DB_UPDATE_FAILED = "DocuMate : Database update failed - {error}"
NOTHING_TO_UPDATE = "DocuMate : No data to update."

# --- Checks that stop the run ----------------------------------------------
NO_DATA = "Found no records to populate.\n\n" + SIGNATURE
NO_STATUS_COLUMN = "'STATUS' column not found. Please check the source file.\n\n" + SIGNATURE
NO_PENDING = "No records found with status 'In Process'.\n\n" + SIGNATURE
MISSING_FIELDS_HEADER = "\nDocuMate : Records with missing mandatory fields detected:\n"
MISSING_FIELDS_ROW = "Row {row} | Missing fields: {fields}"
MISSING_FIELDS = (
    "Warning: Found {count} record(s) with missing values.\n"
    "Please correct those rows in the source file before proceeding.\n\n" + SIGNATURE
)
DUPLICATES = (
    "Warning: Found {count} duplicate {key} entries in pending records.\n"
    "Please check the source file for duplicates.\n\n" + SIGNATURE
)
SORT_SKIPPED = "DocuMate : Some Serial values could not be read as numbers, sorting skipped."

# --- Before anything starts ------------------------------------------------
TEMPLATE_MISSING = "Template not found:\n{path}\n\n" + SIGNATURE
SOURCE_MISSING = "Source file not found:\n{path}\n\n" + SIGNATURE

# --- How it ended ----------------------------------------------------------
SUCCESS = (
    "DocuMate : Mission accomplished!\n"
    "Generated {count} documents successfully{update_text}.\n"
    "{failed_text}"
    "Total time taken: {seconds:.2f} seconds.\n\n"
    "{engine} + {source}.\n" + SIGNATURE
)
RENDER_FAILURES = "{failed} record(s) failed and were left pending - see the console for details.\n"
STATUS_UPDATE_FAILED = (
    "Documents were created and saved to {path}, but {source} was NOT updated: {reason}.\n"
    "Close the file and mark these rows as PRINTED yourself, or they will be generated again on the next run."
)
FILE_LOCKED = (
    "I cannot start because a file I need is open.\n"
    "Please close it and try again.\n\nMission aborted."
)
UNEXPECTED = "DocuMate hit an unexpected error:\n{error}\n\n" + SIGNATURE

# --- Auto-detect loop ------------------------------------------------------
POLL_ENABLED = "DocuMate : auto-detection enabled."
POLL_INTERVAL = "DocuMate : Checking for new records every {seconds} seconds...\n"
POLL_STOPPED = "\nDocuMate : Auto-detection stopped by user."
POLL_ERROR = "DocuMate : Auto-detection error - {error}"


# ---------------------------------------------------------------------------
# Raw-sheet checks
# ---------------------------------------------------------------------------

# A file number that does not start the way this register's numbers do.
PREFIX_MISMATCH = (
    "{count} record(s) have a {column} that does not start with "
    "\"{prefix}\".\n\n"
    "These look like they came from another register, or were typed wrong.\n"
    "Nothing has been generated. Check the rows below, or change the\n"
    "expected prefix in setup/config.py if the numbering has changed."
)

PREFIX_ROW = "    Row {row}: {value}"


# The S No. column feeds the Registration No. printed on the certificate,
# so a blank one cannot simply be ignored.
SERIAL_ASK = (
    "DocuMate : {count} row(s) have no S No.\n"
    "           Number all {total} rows 1-{total} for this run only?\n"
    "           The workbook will not be changed. (yes/no): "
)

SERIAL_RENUMBERED = (
    "DocuMate : Numbered {total} rows 1-{total} for this run."
)

SERIAL_REFUSED = (
    "{count} record(s) have no S No.\n\n"
    "Nothing has been generated. Fill the S No. column in Excel, drag the\n"
    "sequence down to the last record, and run again."
)

SERIAL_UNATTENDED = (
    "{count} record(s) have no S No.\n\n"
    "DocuMate is checking for new records on a timer, so it will not stop\n"
    "to ask whether to number them. Fill the S No. column in Excel and the\n"
    "next check will pick them up."
)


# The sheet itself is missing. The file used to be the only thing checked,
# so a renamed sheet surfaced as a pandas error mid-run instead.
EXCEL_NO_SHEET = (
    "The workbook has no sheet called \"{sheet}\".\n\n"
    "    {path}\n\n"
    "Sheets in this file: {sheets}"
)

# A sheet that is nearly a raw entry sheet, but not quite. Almost always a
# renamed heading. Worth stopping for: the alternative is a certificate
# with half its fields blank, which looks finished and is not.
RAW_SHEET_INCOMPLETE = (
    "This looks like a raw entry sheet, but {found} of {total} expected\n"
    "columns were found.\n\n"
    "Missing: {missing}\n\n"
    "Check the column headings in the sheet. Nothing has been generated."
)


# No signing authority anywhere. Worth its own message: the fix is one
# setting, not an edit to every row, and the generic missing-fields warning
# sends people into the spreadsheet looking for something that was never
# supposed to be typed there.
NO_SIGNING_AUTHORITY = (
    "No signing authority found.\n\n"
    "Every certificate names the officer who signs it. DocuMate will not\n"
    "supply that name itself, so nothing has been generated.\n\n"
    "Add a Signing_Authority column to the sheet and fill it in for the\n"
    "records being printed. It is the one field that has to be chosen\n"
    "rather than worked out, which is why it is the only one left manual."
)