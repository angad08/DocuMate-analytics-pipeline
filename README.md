# DocuMate

Document automation and data validation for consular birth-registration certificates.

DocuMate replaces the manual copying of applicant details from Excel into Word certificates. It started as a Python script and grew into a modular tool supporting Excel, PostgreSQL and Azure SQL, with different document-generation methods for different workloads. An optional Azure Data Factory pipeline loads spreadsheet data into Azure SQL.

![DocuMate workflow](project/diagram/DocuMate%20Workflow%20Architecture%20Diagram.png)

## The problem

Officers prepared certificates by copying applicant details into Word, one record at a time. Excel also held the processing status, so staff had to check missing details, track completed records and work through the backlog manually.

DocuMate brings document generation and status tracking into one workflow, with validation for the Excel versions and a shared database for the database versions and reporting.

## Impact

- **Operational use:** helped the Dispatch team clear approximately 88% of a pending backlog in one week.
- **Benchmark:** DocuMate O processed 1,440 records in about 100 seconds. At the measured manual rate of 315 seconds per record, the equivalent manual workload is an estimated 126 hours.
- **Processing control:** generated records are marked `COMPLETED`, and `PRINTED` once sent to a printer, so subsequent runs select only work that is still outstanding.
- **Reporting:** Power BI can read the same PostgreSQL data used by DocuMate to report on processing volume and backlog.

The benchmark measures automated processing, not the full time needed for review, signing and dispatch.

## How it works

1. **Read pending records** from Excel, PostgreSQL or Azure SQL, depending on the selected version. The Excel versions read the sheet people type into and work the certificate fields out from it; the database versions read columns the database already holds separately.
2. **Apply the version's checks.** Excel versions run the five-stage validation routine, which also checks the columns the certificate fields are built from, so a blank address line or an unreadable date is caught before it prints as a stray comma. Database versions use SQL filtering and database constraints; they do not run that same routine.
3. **Generate certificates** in a combined Word document using Python template rendering or Word Mail Merge. A PDF copy is saved alongside it, converted from the finished Word file so the two always match. PDF export uses Word, so it needs Microsoft Word on Windows; where that is unavailable the Word document is still produced and only the PDF step is skipped.
4. **Update processing status** for successfully generated records, from `IN PROCESS` to `COMPLETED`.
5. **Offer to print** the merged document, and mark what was printed as `PRINTED`. Optional: the files are saved first, so declining loses nothing.
6. **Use the database for reporting**, without copying the records into a separate reporting spreadsheet.

The optional ADF route adds an ingestion step before the database workflow:

```text
Excel in Blob Storage → ADF → Azure SQL → DocuMate → Word certificates
                                            ↓
                                  Processing-status updates
```

ADF loads the records. DocuMate runs separately to generate the documents.

## Versions

| Version | Source | Document generation | Purpose |
| --- | --- | --- | --- |
| **v3** | Excel | `docxtpl` | Production Excel workflow with validation and merged output |
| **X** | Excel | Word Mail Merge | Batches of 250 records for large runs |
| **Y** | Excel | Word Mail Merge | One batch; suitable for smaller runs |
| **Z** | PostgreSQL or Azure SQL | `docxtpl` | Database workflow with polling for pending records |
| **O** | PostgreSQL or Azure SQL | Word Mail Merge | Database workflow with batches of 100 records |

The earlier v1 and v2 scripts established the proof of concept. The five maintained versions share the same modules and select a source and rendering engine through configuration.

In v3, X and Y, validation checks that data and a status column exist, selects pending records, and blocks generation when required fields are missing or file numbers are duplicated. Z and O have this application-level routine disabled and instead use the database's constraints and pending-record queries.

## Optional Azure Data Factory ingestion

![DocuMate Excel-to-database workflow: raw Excel entry, Blob Storage upload, ADF field selection and upserts into related Azure SQL tables](project/diagram/DocuMate-ADF-Ingestion-Workflow.png)

The ADF pipeline reads Excel files from Azure Blob Storage and loads them into related Azure SQL tables. This extends the database workflow; Excel and PostgreSQL versions do not require ADF.

The mapping data flow writes MHA records first, followed by applicant records. It uses file-number keys to insert new records or update matching records. Processing status and issue dates are excluded from the write mapping, preserving those fields on existing records.

State and signing-authority reference data must already be loaded. The Python loader is a separate option: it populates all four tables and skips existing records, whereas ADF writes to the two tables below and can update matching rows.

Definitions are stored in [`adf/`](adf/) through Data Factory's Git integration. Saving definitions in Git and publishing them to a running factory are separate steps.

| Definition | Purpose |
| --- | --- |
| [Pipeline](adf/pipeline/DocuMateADF.json) | Runs the mapping data flow |
| [Data flow](adf/dataflow/DocumateDataFlow.json) | Selects columns and inserts or updates the two tables |
| [Excel dataset](adf/dataset/Documate_Excel.json) | Reads the workbook from Blob Storage |
| [Applicant dataset](adf/dataset/DS_Documate_DB.json) | Connects to the applicant table |
| [MHA dataset](adf/dataset/DS_Documate_MHA.json) | Connects to the MHA table |
| [Linked services](adf/linkedService/) | Blob Storage and Azure SQL connections |

Reconfigure storage, database connections and authentication for your environment before running the pipeline. The repository also contains a self-hosted runtime definition for local-file access; the committed Excel source uses Blob Storage.

## Quick start

Install the dependencies:

```bash
pip install -r requirements.txt
```

Check the installation and list the available versions:

```bash
python main.py --check
python main.py --list
```

Run the Excel-based production workflow:

```bash
python main.py v3
```

Each run asks for the Excel file and Word template with a file picker, unless `DOCUMATE_EXCEL_PATH` and the template setting are set in `config.env`. The picker opens in the folder used last time, and those choices are kept in `files/last_run.json`, which is ignored by Git. The Excel versions are `v3`, `x` and `y`. The repository ships one worked sample, `files/data/SAMPLE_CASE_STUDY.xlsx` with `files/templates/SAMPLE_CASE_STUDY.docx`; birth-register workbooks and templates are not included. The Mail Merge versions (`x`, `y`, `o`) require Microsoft Word on Windows. Generated documents go to `files/output/`, which is intentionally ignored by Git.

Your own workbook needs a `Signing_Authority` column before a run will start; see [The Excel sheet](#the-excel-sheet).

### Running without Python

Each version can be built as a single `.exe` that carries its own Python and
every library it needs, so the PC it runs on needs neither installed:

```bash
pip install pyinstaller
python packaging\build_exe.py          # all five
python packaging\build_exe.py v3 x     # or just some
```

The exes are written to `dist\`. Keep them there: an exe looks one folder up
for `files\` and the `config.env` file, so one moved elsewhere will not find the
templates or the workbook.

Microsoft Word is still required for the Mail Merge versions, for PDF export
and for printing - those drive Word itself rather than reimplementing it.

The window icon comes from `logo\documate.ico`. Windows will not accept a PNG
however it is named, so `packaging\make_icon.py` builds the `.ico` from
`logo\documate.ico.png` whenever the logo changes.

## The birth register sheet

This section is about the birth register only. DocuMate recognises its
sheet by the column headings; any other sheet's columns go to the template
as they are - see [Making any kind of document](project/docs/CASE_STUDY.md).

The birth register's sheet is the one people type into. DocuMate works the
certificate fields out from it: names in capitals, `F` as `FEMALE`, the date
and place of birth joined, and the registration number built from the serial
and the year the record was registered.

Those rules used to be Excel formulas on a second sheet that had to be
maintained by hand. They are in [`sources/excel_format.py`](sources/excel_format.py)
now, so the workbook holds data and the repository holds the rules. A workbook
that still has a prepared sheet keeps working: point `DOCUMATE_SHEET_NAME` at
it and it is used as it is, because the shape is recognised from the column
names rather than the sheet name.

The sheet needs `S No.`, `File No`, `Name`, `Sex`, `Date of birth`,
`Place of birth`, both parents, `Registration date`, `ADD 1`, `MHA_File_No`,
`MHA_Date` and `Signing_Authority`. `ADD 2` and `ADD 3` are optional and are
left out of the address when empty. Dates are read day-first, so `03/04/2025`
is the 3rd of April, and a format that cannot be read without guessing stops
the run rather than being interpreted.

**The signing authority is the one field DocuMate will not supply.** It names
the officer on an official document, so it comes from the sheet or the run
stops. Different records can name different officers.

Two optional checks are off by default, because the right value differs
between registers. Set either in `config.env` and a record numbered elsewhere
stops the run:

```
DOCUMATE_FILE_PREFIX=IND/MEL/
DOCUMATE_MHA_PREFIX=
```

## Printing

After the documents are generated, DocuMate offers to send the merged file to
a printer. It is optional: the Word file and its PDF are saved first, so
declining costs nothing and the file can be printed by hand later.

The printer is chosen rather than assumed. Pressing Enter takes the one the
machine already prints to, so the usual answer is a single keypress, but the
list is shown because a Windows default is often not a printer at all.
Printers that write a file rather than paper - Print to PDF, XPS, OneNote,
Fax - are marked as such, since sending a hundred certificates to one of
those produces nothing and looks like it worked.

The printer is checked before the job is sent. An offline, jammed or
unreachable printer stops the run with a message naming the problem, rather
than a job disappearing into a queue nobody is watching.

The `.docx` is printed, not the PDF. Word opens a PDF by converting it back
into an editable document, which reflows the layout; and handing the file to
whatever program owns `.pdf` prints to the default printer with no way to
choose another. The PDF is made by Word's own export, so printing the Word
file puts the same pages on paper with the printer still under control.

Set `DOCUMATE_PRINTING=0` to turn the offer off, for an office that prints
by hand or for an unattended run with nobody to answer the question.

### What the statuses mean

| Status | Meaning |
| --- | --- |
| `IN PROCESS` | Not yet done. Blank or an unrecognised value counts as this |
| `COMPLETED` | The document has been generated |
| `PRINTED` | The job was accepted by a printer |

A run that generates and is not printed stops at `COMPLETED`, which is what
happened. Both `COMPLETED` and `PRINTED` count as finished, so neither is
picked up again on the next run.

`PRINTED` means the job was accepted and spooled without error. That is not
proof that paper exists - a job can still be cancelled at the device - but it
is the strongest thing that can be known without asking somebody to go and
count the stack.

A database needs its `status` constraint to allow the middle value before the
database versions will write it:

```sql
ALTER TABLE applicant DROP CONSTRAINT chk_applicant_status;
ALTER TABLE applicant ADD CONSTRAINT chk_applicant_status
  CHECK (UPPER(status) IN ('IN PROCESS', 'COMPLETED', 'PRINTED'));
```

## Modular structure

```text

├── main.py          command-line runner
├── flow/            validation, pipeline, status updates
├── sources/         Excel, Azure SQL and PostgreSQL readers
├── engines/         docxtpl and Word Mail Merge renderers
├── versions/        v3, X, Y, Z and O settings
├── setup/           paths, configuration, file picker, messages, UI
├── files/           synthetic data, templates, generated output
├── adf/             Azure Data Factory pipeline, dataflow and datasets
├── packaging/       standalone .exe builds
├── tools/           database seeding
├── tests/           pytest suite
└── project/         documentation, schema and workflow diagram
```

To change a version, edit [the version registry](versions/registry.py). To change the workflow shared by every version, edit [the pipeline](flow/pipeline.py). Connection and file-path settings belong in [configuration](setup/config.py), while private values belong in `config.env`.

## Database setup

Z and O use Azure SQL or PostgreSQL, chosen by `DOCUMATE_DB_BACKEND`. Create the
tables with the script for your database:

```bash
# Azure SQL (default)
sqlcmd -S YOUR_SERVER.database.windows.net -d YOUR_DB -U YOUR_USER -P YOUR_PASSWORD        -i project/database_schema/DocuMate_Data_Schema_AzureSQL.sql

# PostgreSQL
psql -h YOUR_HOST -U YOUR_USER -d YOUR_DB -f project/database_schema/DocuMate_Data_Schema.sql
```

Enter the database connection values in `config.env`, the one settings file in the project folder. It is ignored by Git; never place passwords or connection credentials in source files.

```text
DOCUMATE_DB_BACKEND=azuresql
DOCUMATE_DB_HOST=your-server.database.windows.net
DOCUMATE_DB_NAME=your-database
DOCUMATE_DB_USER=your-user
DOCUMATE_DB_PASSWORD=your-password
```

Leave `DOCUMATE_DB_PORT` unset and the backend's own port is used: 1433 for Azure
SQL, 5432 for PostgreSQL. Azure SQL also needs Microsoft ODBC Driver 18 installed
and your client IP allowed under the SQL server's Security -> Networking. Run
`python main.py --check` to see what was picked up.

Load records with `python -m tools.seed_database --insert --file YOUR_WORKBOOK.xlsx`;
it writes to whichever backend is selected. The sample workbook it used by default
is no longer in the repository, so pass `--file`.

## Any kind of document

DocuMate is not tied to any one document. Every version asks for the Excel
file and the Word template when it starts - unless they are filled in
`config.env`, the one settings file - fills each `{{ Column_Name }}` in the
template from that column, and asks which column identifies a row. Nothing
to set up per document - see
[Making any kind of document](project/docs/CASE_STUDY.md). A worked sample
is in `files/`: `SAMPLE_CASE_STUDY.xlsx` and `SAMPLE_CASE_STUDY.docx`.

## Further reading

- [Making any kind of document](project/docs/CASE_STUDY.md)
- [Evolution and benchmarks](project/docs/EVOLUTION.md)
- [Template setup](project/docs/TEMPLATES.md)
- [Engineering decisions](project/docs/DESIGN_DECISIONS.md)
- [Modularisation map](project/docs/MODULARISATION.md)

## Author

**Angad Kadam**

[LinkedIn](https://linkedin.com/in/angad-kadam-03b606159) · [GitHub](https://github.com/angad08)
