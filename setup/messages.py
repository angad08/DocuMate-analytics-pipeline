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
ASK_UPDATE = "DocuMate : Do you want me to mark the statuses as PRINTED for {count} records? (yes/no): "
UPDATING = "\nDocuMate : Updating {target} statuses...\n"
UPDATE_SUFFIX = " and updated {target} in {seconds:.2f} seconds"
EXCEL_UPDATED = "DocuMate : Sheet '{sheet}' updated successfully on {date} at {clock}."
EXCEL_NO_STATUS_COLUMN = "DocuMate : STATUS column not found - skipping Excel update."
EXCEL_LOCKED = "DocuMate : Cannot update Excel while the file is open. Close it and retry."
EXCEL_UPDATE_ERROR = "DocuMate : Error updating Excel - {error}"
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
    "Total time taken: {seconds:.2f} seconds.\n\n"
    "{engine} + {source}.\n" + SIGNATURE
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
