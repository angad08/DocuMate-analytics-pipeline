# How DocuMate evolved

Seven versions, built one after another. Each one answered a problem the
previous one had. This is the long version of the story the README summarises.

[Back to the README](../README.md)

---

## How the System Evolved

The project went through several iterations. Each version was a response to a specific problem discovered in the previous one.

### Chronology

```text
v1 -> v2 -> v3 -> (Y -> X) + (Z) -> O
```

- **v1 to v3** built the original automation foundation
- **Y and X** improved how documents are rendered (Mail Merge branch)
- **Z** improved how data is stored and tracked (database branch)
- **O** brought both branches together into one system

### v1.0 -- Proof of concept

The first version was a flat script. It read an Excel file, selected every row not yet marked "PRINTED" (any record whose status is `IN PROCESS`, blank, or anything other than `PRINTED`), and generated one Word document per record using `docxtpl`. No error handling, no validation, no structure. Its only job was to prove the idea worked: that Python could turn Excel rows into Word documents faster than a human copying and pasting.

It worked. That was enough to justify building further.

### v2.0 -- Stabilisation

v1 was fragile. If anything went wrong, it crashed silently. v2 wrapped the same logic in a proper class structure, added error handling, Windows popup messages for user feedback, and portable file paths so the tool could run on different machines without editing hardcoded directories. It also added standalone `.exe` packaging support via PyInstaller so the end user did not need Python installed.

### v3.0 / PLUS -- Data reliability

v2 worked, but it trusted its input. If someone left a field blank or entered a duplicate file number, the system would generate a bad certificate without warning. v3 added a five-stage validation pipeline that checks the data before any document is generated: dataset existence, status column verification, active record filtering, mandatory field validation, and duplicate detection. The active record filter selects every row not yet marked "PRINTED" (`IN PROCESS`, blank, or any other non-`PRINTED` status), so records with a missing or non-standard status are no longer silently skipped. If any check fails, processing stops immediately.

v3 also introduced merged document output (all certificates combined into one print-ready file using `docxcompose`) and parallel processing using `ProcessPoolExecutor` to render documents simultaneously across multiple CPU cores.

### Y -- Rendering engine swap

Up to this point, every version used `docxtpl` as the rendering engine: Python fills a Word template using Jinja-style placeholders, one document at a time. `docxtpl` is solid. It is cross-platform, Python-native, requires no external software, and works anywhere Python runs, including servers and Linux environments.

But the workplace already had Microsoft Word installed on every machine. Word has its own built-in document generation engine: Mail Merge. It is designed for exactly this job -- reading a data source and producing filled documents in bulk. The question was: could Python control Word's Mail Merge engine directly instead of rendering documents itself?

Y answered that question. Using `win32com` and COM automation, Python opens the Word template, connects the data source, sets the record range, and tells Word to execute the merge. Python stops being the renderer and becomes the orchestrator. Word does what it was built to do, and Python controls when, what, and how many.

One practical detail applies to all three Mail Merge versions. Pointing Word at an `.xlsx` means connecting through the ACE OLEDB Excel provider, which treats the workbook as a database where each sheet is a table -- so Word raises its "Select Table" dialog and waits for a human before it will run the query, making unattended runs impossible. It does this regardless of the `SQLStatement` passed, and even when the workbook contains exactly one sheet. Y, X and O therefore write the rows to process to a temporary CSV and merge from that: a CSV has one implicit table, so there is nothing to choose and nothing to ask.

This was the first crossover: Python driving a desktop application programmatically to automate a process that previously required manual clicks through Word's UI.

Like the other Excel versions, Y selects every row not yet marked "PRINTED" (`IN PROCESS`, blank, or any other non-`PRINTED` status) when building the record range it hands to Word, and the same condition guards the write-back step that marks rows as `PRINTED` afterward.

### X -- Batched execution

Y worked well on small volumes but hung on large ones. When Word tried to process 1,000+ records in a single `Execute()` call, the COM interface became unresponsive. X solved this by splitting the record range into batches of 250. Each batch produces a temporary `.docx` file, and after all batches complete, the files are combined into one final document using Word's `InsertFile` method. Temp files are cleaned up automatically.

X keeps the same Excel record selection as the other Excel versions -- every row not yet marked "PRINTED" (`IN PROCESS`, blank, or any other non-`PRINTED` status).

Those filtered rows do not have to sit together in the sheet. X writes them to a temporary CSV and merges from that, so the records are always numbered 1..N with no gaps, however scattered they are in the source. Earlier revisions merged against the full worksheet, which meant a `From`/`To` range spanning a gap would have swept in already-`PRINTED` rows -- so the run was refused outright unless the selection happened to be contiguous. Filtering before Word rather than inside it removes that restriction entirely.

This turned a hanging system into the fastest version of DocuMate at the time.

### Z (v4.0) -- Data layer migration

By this point, Excel was the bottleneck. It was acting as both the data source and the status tracker, which made it difficult to connect reporting tools or run concurrent operations. Z replaced the entire data layer with PostgreSQL. Records are pulled via SQL joins across normalised tables (`applicant`, `ministryofhomeaffairs`, `ib_authority`), and status updates go directly to the database. Because the data now lives in a proper database, Power BI can connect to the same source for live dashboards without any extra integration work.

Z also added an auto-detection mode: a polling loop that checks for new "IN PROCESS" records at a configurable interval, turning DocuMate from a run-once script into a lightweight background service.

### O -- The latest convergence version

O is the convergence. It takes Z's PostgreSQL backend and combines it with X's batched Word Mail Merge engine.

The challenge was that Word's Mail Merge cannot connect to a database directly -- it expects a flat file (Excel, CSV, or Access). O bridges this gap: records are pulled from PostgreSQL via `load_data()`, written to a temporary CSV file, and that temp file becomes the Mail Merge data source. After processing, the temp file is cleaned up automatically. The database handles data, Word handles rendering, and Python orchestrates everything in between.

O closes the loop:

**data source -> validation -> generation -> write-back -> reporting**

O is not the fastest version in raw time (v2 is faster on small batches because it has no COM overhead). But it is the only version that combines the PostgreSQL backend, Word Mail Merge engine, batch execution, database status updates, auto-detection mode, and Power BI integration into one system. That is why it is called O -- it is the complete version.

---

## Two Rendering Engines

DocuMate supports two rendering engines, each with its own strengths.

**Python rendering (docxtpl)** is used in v1, v2, v3, and Z. It fills Word templates using Jinja-style placeholders entirely in Python. It is cross-platform, runs on any operating system, requires no external software, and works well in server or headless environments. The parallel processing in v3 and Z uses `ProcessPoolExecutor` to render documents across multiple CPU cores simultaneously, with `docxcompose` merging the results. This engine is the right choice when Word is not available or when the system needs to run outside of Windows.

**Word Mail Merge (COM automation)** is used in Y, X, and O. Python controls Microsoft Word's native Mail Merge engine through `win32com`. Instead of rendering documents itself, Python sets the data source, record range, and batch size, then tells Word to execute. Word produces the documents using its own rendering pipeline. This approach is faster on large volumes because Word processes records internally rather than through Python's per-document loop. It requires Microsoft Word installed and runs on Windows only.

From a business perspective, both paths achieve the same outcome: automated, bulk-generated, print-ready documents. The choice depends on the environment and requirements.
