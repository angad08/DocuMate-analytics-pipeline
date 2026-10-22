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
# A bar rather than a bare counter: on 1,440 records a number alone says
# nothing about how much is left. {bar} is filled in by setup.ui.progress_bar.
RECORD_PROGRESS = "DocuMate : [{bar}] {percent:3d}%  {done}/{total} records"
CREATING_FOLDER = "\nDocuMate : Output folder '{folder}' does not exist. Creating it..."
SAVED = (
    "\nDocuMate : Generated and merged {count} documents (in ID order)"
    " in {seconds:.2f} seconds and saved in:\n{path}"
)
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

# --- Marking records as completed --------------------------------------------
ASK_UPDATE = "DocuMate : Mark {count} as completed?{failed} (yes/no): "
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
MISSING_FIELDS_HEADER = "\nDocuMate : Rows that need fixing before documents can be made:\n"
# Superseded by the table in flow/checks.problem_table; kept so anything still
# importing it does not break.
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
    "Close the file and mark these rows as COMPLETED yourself, or they will be generated again on the next run."
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

# --- Choosing the files (setup/picker.py) ---------------------------------
# The template picked was made for the other engine. Common in a file
# picker - BIRTH_REG_FORMAT.docx and BIRTH_REG_FORMAT_MM.docx sit side by
# side - and the wrong one prints documents with every value missing.
WRONG_TEMPLATE_KIND = {
    "docxtpl": (
        "{template} is a Word Mail Merge template, but {version} fills\n"
        "{{{{ placeholder }}}} templates.\n\n"
        "Pick the template without _MM, or run X or Y to use this one."
    ),
    "mailmerge": (
        "{template} is a {{{{ placeholder }}}} template, but {version} uses\n"
        "Word Mail Merge.\n\n"
        "Pick the _MM template, or run v3 to use this one."
    ),
}
TYPE_PATH = "DocuMate : {title} - type or paste its path: "
NOTHING_CHOSEN = "No file or column was chosen, so nothing has been generated."
CHOOSE_SHEET_HEADING = "\nDocuMate : {workbook} has more than one sheet:"
CHOOSE_SHEET = "Sheet to read"
CHOOSE_ID_HEADING = (
    "\nDocuMate : Which column identifies each row? It must be filled in and\n"
    "           different on every row - each row is marked done by it."
)
CHOOSE_ID = "ID column"
RUN_FILES = (
    "DocuMate : Excel {excel} (sheet {sheet}), template {template}, "
    "rows identified by {id}.\n"
)
RUN_TEMPLATE = "DocuMate : Template {template}.\n"

# The template asks for something the sheet does not have. Usually a typo
# in a placeholder or a renamed heading; either way every document would
# print with that gap, so it stops.
TEMPLATE_FIELDS_MISSING = (
    "The template asks for {count} field(s) the sheet does not have:\n\n"
    "    {fields}\n\n"
    "Each placeholder must match a column heading - \"Passport No\" is\n"
    "{{{{ Passport_No }}}}. Columns in the sheet:\n\n"
    "    {columns}\n\n"
    "Fix the placeholder or the heading. Nothing has been generated."
)

# Blank cells in columns the template prints. Sometimes that is right - a
# Remarks column - so someone at the keyboard is asked rather than refused.
BLANKS_ASK = (
    "DocuMate : {count} row(s) have blank cells in fields the template prints\n"
    "           (listed above). Generate anyway, leaving them blank? (yes/no): "
)

# A case's sheet without the column that tells its rows apart. Caught on
# read, because the documents would otherwise be made and only the status
# write-back would fail.
NO_ID_COLUMN = (
    "The sheet has no \"{column}\" column.\n\n"
    "DocuMate uses it to tell the rows apart and to mark each one done, so\n"
    "every row needs a different value in it. Add the column, or set\n"
    "DOCUMATE_ID_COLUMN in config.env to one the sheet already has.\n\n"
    "Columns in this sheet: {columns}\n\n"
    "Nothing has been generated."
)
EXCEL_NO_ID_COLUMN = "DocuMate : {column} column not found - cannot update Excel."
DB_NO_ID_COLUMN = (
    "Table {table} has no column called {column}, so the rows cannot be\n"
    "marked done. Set DOCUMATE_DB_ID_COLUMN in config.env to the table's ID column."
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

# ---------------------------------------------------------------------------
# Printing
# ---------------------------------------------------------------------------

ASK_PRINT = "DocuMate : Print the {count} merged certificate(s)? (yes/no): "

CHOOSE_PRINTER = "Printer"

CHOOSE_AGAIN = "   Type a number from 1 to {count}, or n to skip."

PRINTERS_HEADING = "\nPrinters:"

# Marked so a virtual printer is obvious before it is chosen rather than
# after a hundred certificates have gone into a file nobody asked for.
PRINTER_DEFAULT_TAG = "   [default]"
PRINTER_VIRTUAL_TAG = "  (writes a file, not paper)"

NO_PRINTERS = (
    "DocuMate : No printers are set up on this PC, so there is nothing to\n"
    "           print to. The documents are saved and can be printed later."
)

PRINTING_UNAVAILABLE = (
    "Printing needs pywin32, which could not be loaded: {error}"
)

PRINTER_UNREACHABLE = (
    "The printer \"{printer}\" could not be reached: {error}"
)

PRINTER_NOT_READY = (
    "The printer \"{printer}\" is {problem}."
)

PRINT_FILE_MISSING = (
    "The merged document is no longer at:\n\n    {path}"
)

PRINT_FAILED = (
    "The job was not accepted by \"{printer}\":\n\n    {error}"
)

PRINT_SENDING = "DocuMate : Sending {count} certificate(s) to {printer}..."

PRINT_SENT = "DocuMate : Print job accepted."

# Said plainly, because the status it writes claims more than it can prove.
PRINT_NOT_CONFIRMED = (
    "DocuMate : The job was accepted by the printer. Check the stack before\n"
    "           treating these as issued."
)

PRINT_SKIPPED = (
    "DocuMate : Nothing was printed. The documents are saved in:\n\n    {path}"
)

PRINT_PROBLEM = (
    "{problem}\n\n"
    "Nothing has been printed and no statuses were changed. The documents\n"
    "are saved and can be printed again once the printer is ready."
)
