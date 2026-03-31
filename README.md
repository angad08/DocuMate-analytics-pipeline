# DocuMate

DocuMate is a document automation and data validation system I built to replace a slow, repetitive, error-prone manual workflow that relied heavily on Excel copy-paste operations.

What started as a simple Python script gradually evolved into a structured system that combines:

- automated document generation
- multi-stage data validation
- Excel and PostgreSQL data sources
- Python rendering and COM-controlled Word Mail Merge
- processing status updates
- Power BI reporting

Manual processing of roughly **1,440 records typically required nearly two months of manual effort**.
DocuMate reduces that to **minutes of automated processing**, with safeguards that stop bad data before it turns into bad output.

---

## Architecture

![DocuMate Workflow Architecture](diagram/DocuMate%20Workflow%20Architecture%20Diagram.png)

---

## The Problem

The original workflow involved officers manually copying data from Excel into Word templates to generate certificates.

Each record required:

- Opening a Word template
- Copying values from Excel fields
- Pasting them into the correct placeholders
- Formatting fields manually
- Saving the document
- Repeating the same process for every record

This process had several issues:

- Extremely time-consuming
- High chance of human error
- No systematic validation of incoming data
- No structured storage for records
- No visibility into processing volumes or trends

The workflow was effectively Excel acting as both a database and a processing system, which made it difficult to scale or monitor.

The core business problem was not just speed. It was also consistency, reliability, traceability, and operational control. DocuMate was built to solve those problems, not just to "make Word files faster."

---

## What I Built

DocuMate automates the full document processing lifecycle:

1. Reads applicant records from **Excel or PostgreSQL**
2. Runs a **multi-stage data validation process**
3. Maps validated data into **Word document templates**
4. Generates certificates using **Python rendering (docxtpl)** or **Word's native Mail Merge engine via COM automation**
5. Merges all documents into a **single print-ready file**
6. Updates processing status in Excel or the database
7. Exposes the same data to **Power BI dashboards for reporting**

Instead of manually generating hundreds or thousands of files, the system processes records in bulk while enforcing validation rules.

---

## Business Outcome

DocuMate improves the process in five key ways:

### 1. Speed
Bulk generation reduces turnaround time dramatically compared with manual handling.

### 2. Accuracy
Validation checks catch missing fields, bad records, and duplicates before output is created.

### 3. Consistency
Every document follows the same generation logic and formatting pipeline.

### 4. Control
Processing status is written back to Excel or the database after successful generation.

### 5. Visibility
With PostgreSQL as the data layer, the same records feed Power BI dashboards for operational reporting.

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

The first version was a flat script. It read an Excel file, filtered rows marked "IN PROCESS", and generated one Word document per record using `docxtpl`. No error handling, no validation, no structure. Its only job was to prove the idea worked: that Python could turn Excel rows into Word documents faster than a human copying and pasting.

It worked. That was enough to justify building further.

### v2.0 -- Stabilisation

v1 was fragile. If anything went wrong, it crashed silently. v2 wrapped the same logic in a proper class structure, added error handling, Windows popup messages for user feedback, and portable file paths so the tool could run on different machines without editing hardcoded directories. It also added standalone `.exe` packaging support via PyInstaller so the end user did not need Python installed.

### v3.0 / PLUS -- Data reliability

v2 worked, but it trusted its input. If someone left a field blank or entered a duplicate file number, the system would generate a bad certificate without warning. v3 added a five-stage validation pipeline that checks the data before any document is generated: dataset existence, status column verification, active record filtering, mandatory field validation, and duplicate detection. If any check fails, processing stops immediately.

v3 also introduced merged document output (all certificates combined into one print-ready file using `docxcompose`) and parallel processing using `ProcessPoolExecutor` to render documents simultaneously across multiple CPU cores.

### Y -- Rendering engine swap

Up to this point, every version used `docxtpl` as the rendering engine: Python fills a Word template using Jinja-style placeholders, one document at a time. `docxtpl` is solid. It is cross-platform, Python-native, requires no external software, and works anywhere Python runs, including servers and Linux environments.

But the workplace already had Microsoft Word installed on every machine. Word has its own built-in document generation engine: Mail Merge. It is designed for exactly this job -- reading a data source and producing filled documents in bulk. The question was: could Python control Word's Mail Merge engine directly instead of rendering documents itself?

Y answered that question. Using `win32com` and COM automation, Python opens the Word template, connects the Excel data source, sets the record range, and tells Word to execute the merge. Python stops being the renderer and becomes the orchestrator. Word does what it was built to do, and Python controls when, what, and how many.

This was the first crossover: Python driving a desktop application programmatically to automate a process that previously required manual clicks through Word's UI.

### X -- Batched execution

Y worked well on small volumes but hung on large ones. When Word tried to process 1,000+ records in a single `Execute()` call, the COM interface became unresponsive. X solved this by splitting the record range into batches of 250. Each batch produces a temporary `.docx` file, and after all batches complete, the files are combined into one final document using Word's `InsertFile` method. Temp files are cleaned up automatically.

This turned a hanging system into the fastest version of DocuMate at the time.

### Z (v4.0) -- Data layer migration

By this point, Excel was the bottleneck. It was acting as both the data source and the status tracker, which made it difficult to connect reporting tools or run concurrent operations. Z replaced the entire data layer with PostgreSQL. Records are pulled via SQL joins across normalised tables (`applicant`, `ministryofhomeaffairs`, `ib_authority`), and status updates go directly to the database. Because the data now lives in a proper database, Power BI can connect to the same source for live dashboards without any extra integration work.

Z also added an auto-detection mode: a polling loop that checks for new "IN PROCESS" records at a configurable interval, turning DocuMate from a run-once script into a lightweight background service.

### O -- The latest convergence version

O is the convergence. It takes Z's PostgreSQL backend and combines it with X's batched Word Mail Merge engine.

The challenge was that Word's Mail Merge cannot connect to a database directly -- it expects a flat file (Excel, CSV, or Access). O bridges this gap: records are pulled from PostgreSQL via `load_data()`, written to a temporary Excel file, and that temp file becomes the Mail Merge data source. After processing, the temp file is cleaned up automatically. The database handles data, Word handles rendering, and Python orchestrates everything in between.

O closes the loop:

**data source -> validation -> generation -> write-back -> reporting**

O is not the fastest version in raw time (v2 is faster on small batches because it has no COM overhead). But it is the only version that combines the PostgreSQL backend, Word Mail Merge engine, batch execution, database status updates, auto-detection mode, and Power BI integration into one system. That is why it is called O -- it is the complete version.

---

## Two Rendering Engines

DocuMate supports two rendering engines, each with its own strengths.

**Python rendering (docxtpl)** is used in v1, v2, v3, and Z. It fills Word templates using Jinja-style placeholders entirely in Python. It is cross-platform, runs on any operating system, requires no external software, and works well in server or headless environments. The parallel processing in v3 and Z uses `ProcessPoolExecutor` to render documents across multiple CPU cores simultaneously, with `docxcompose` merging the results. This engine is the right choice when Word is not available or when the system needs to run outside of Windows.

**Word Mail Merge (COM automation)** is used in Y, X, and O. Python controls Microsoft Word's native Mail Merge engine through `win32com`. Instead of rendering documents itself, Python sets the data source, record range, and batch size, then tells Word to execute. Word produces the documents using its own rendering pipeline. This approach is faster on large volumes because Word processes records internally rather than through Python's per-document loop. It requires Microsoft Word installed and runs on Windows only.

From a business perspective, both paths achieve the same outcome: automated, bulk-generated, print-ready documents. The choice depends on the environment and requirements.

---

## Repository Structure

```
DocuMate/
├── README.md
├── requirements.txt
├── .gitignore
│
├── src/
│   ├── DocuMate_v1_0.py
│   ├── DocuMate_v2_0.py
│   ├── DocuMatePLUS_v3_0.py
│   ├── DocuMateY.py
│   ├── DocuMateX.py
│   ├── DocuMateZ.py
│   ├── DocuMateO.py
│   └── loadData.py
│
├── sql/
│   └── DocuMate_Data_Schema.sql
│
├── templates/
│   ├── DOCUMENT_TEMPLATE_FILE.docx
│   └── DOCUMENT_TEMPLATE_FILE_MM.docx
│
├── data/
│   ├── DocuMate_DataFrame.xlsx
│   └── test_data/
│       └── Test_Insert_data.xlsx
│
├── dashboard/
│   └── CGI-MELB.pbix
│
├── diagram/
│   └── DocuMate Workflow Architecture Diagram.png
│
└── output_files/
```

---

## Data Validation Pipeline

One of the biggest improvements introduced in later versions was a validation stage that runs before document generation.

DocuMate performs five checks:

1. **Dataset existence check** -- Confirms that data is available before processing begins.
2. **Status column verification** -- Ensures required fields exist in the dataset.
3. **Active record filtering** -- Filters only records marked "IN PROCESS".
4. **Mandatory field validation** -- Identifies rows with missing values and reports the exact fields.
5. **Duplicate detection** -- Prevents duplicate file numbers from entering the system.

If any validation step fails, processing stops immediately and the issue is reported.

This protects the process, not just the code.

---

## PostgreSQL Architecture

The later versions of DocuMate migrated the data layer to PostgreSQL to move away from Excel acting as a database.

The schema is included in the repository:

```
sql/DocuMate_Data_Schema.sql
```

Running the schema creates the required normalised tables:

- `applicant`
- `ministryofhomeaffairs`
- `ib_authority`
- `state`

This structure separates core record data, authority information, and state mapping while allowing SQL joins during document generation.

Because the data now lives in a proper database, Power BI can connect to the same source for live dashboards. DocuMate handles generation and status updates. Power BI reads the same tables for reporting and monitoring. No custom Python-to-Power-BI integration is required.

The demonstration environment uses Supabase (managed PostgreSQL), but DocuMate works with any PostgreSQL-compatible host.

---

## Database Setup

Create a PostgreSQL database and run the included schema file:

```bash
psql -h YOUR_HOST -U YOUR_USER -d YOUR_DB -f sql/DocuMate_Data_Schema.sql
```

Update the database credentials in:

```
src/DocuMateZ.py
src/DocuMateO.py
src/loadData.py
```

Example configuration:

```python
db_config = {
    "host": "YOUR_DB_HOST",
    "database": "YOUR_DB_NAME",
    "user": "YOUR_DB_USER",
    "password": "YOUR_DB_PASSWORD",
    "port": 5432,
    "sslmode": "require"
}
```

To insert sample records:

```bash
python src/loadData.py
```

This loads test data from `data/test_data/Test_Insert_data.xlsx`.

---

## Running the System

### Excel + Python rendering (v1, v2, v3)

```bash
python src/DocuMate_v1_0.py
python src/DocuMate_v2_0.py
python src/DocuMatePLUS_v3_0.py
```

These versions read records from `data/DocuMate_DataFrame.xlsx` and render documents using `docxtpl`.

### Excel + Word Mail Merge (Y, X)

```bash
python src/DocuMateY.py
python src/DocuMateX.py
```

These versions use Word's native Mail Merge engine via COM automation. They read from `data/DocuMate_DataFrame.xlsx`. Y processes all records in a single execution. X splits into batches of 250 to handle large volumes without Word hanging.

Requires Microsoft Word desktop installed. Windows only.

### Database + Python rendering (Z)

```bash
python src/DocuMateZ.py
```

Reads from PostgreSQL, renders using `docxtpl`. Supports **auto-detection mode**, polling the database for new "IN PROCESS" records at a configurable interval.

### Database + Word Mail Merge (O)

```bash
python src/DocuMateO.py
```

The latest convergence version. Reads from PostgreSQL, writes a temporary Excel bridge, processes via batched Word Mail Merge, cleans up the temp file automatically. Supports auto-detection mode, same as Z.

---

## Template Setup

DocuMate uses two types of Word templates depending on the rendering engine. Both live in the `templates/` folder.

### docxtpl template (v1, v2, v3, Z)

`DOCUMENT_TEMPLATE_FILE.docx` is a standard Word document with Jinja-style placeholders. These are double-curly-brace tags that `docxtpl` replaces with data at render time.

To create or edit the template, open a `.docx` file in Microsoft Word and type the placeholders directly where each value should appear. For example:

```
Name: {{Name}}
Date of Birth: {{When_and_where_born}}
Serial: {{Serial}}
```

No special Word configuration is needed. The file is a normal `.docx` document -- `docxtpl` reads the placeholders and fills them in using Python. This works on any operating system without Microsoft Word installed at runtime.

### Word Mail Merge template (Y, X, O)

`DOCUMENT_TEMPLATE_FILE_MM.docx` is a Word Mail Merge template. Unlike the docxtpl template, this one requires a one-time configuration inside Microsoft Word to link it to a data source and insert merge fields.

#### Steps to create the template

1. Open the base Word template in **Microsoft Word**.
2. Go to the **Mailings** tab in the ribbon.
3. Click **Start Mail Merge** and select **Letters** (or **Normal Word Document**).
4. Click **Select Recipients** and choose **Use an Existing List**.
5. Browse to `data/DocuMate_DataFrame.xlsx` and select it.
6. If prompted, select the **DocuMateSRC** sheet.
7. Word now knows which data source the template is linked to.
8. Place your cursor where each field should appear in the document and click **Insert Merge Field** to add the placeholders.

### Available fields

Both templates use the same field names, which correspond to the column names produced by DocuMate's data pipeline:

- `File_Number`
- `Serial`
- `Name`
- `Sex`
- `When_and_where_born`
- `Name_of_the_Father`
- `Name_of_the_Mother`
- `Description_and_residence_of_informant`
- `Registration_date`
- `MHA_File_And_date`
- `Signing_Authority_Name`
- `Signing_Authority_Name_Designation`

In the docxtpl template, these appear as `{{Name}}`, `{{Serial}}`, etc. In the Mail Merge template, these are inserted via Word's **Insert Merge Field** button.

9. Once all merge fields are placed, save the Mail Merge template as `DOCUMENT_TEMPLATE_FILE_MM.docx` in the `templates/` folder.

#### Note for DocuMateO

DocuMateO does not use the Excel file as its data source at runtime. It pulls records from PostgreSQL and writes them to a temporary Excel file as a bridge for Word's Mail Merge engine. The template still needs to be configured once with the Excel file (steps above) so that Word recognises the merge fields. At runtime, DocuMateO overrides the data source connection to point at the temp file automatically.

---

## Power BI Integration

The Power BI dashboard connects to the same PostgreSQL database used by DocuMate.

DocuMate handles generation and status updates. Power BI reads the same tables for reporting and monitoring.

Typical dashboard insights include:

- Application volume trends
- Processing backlog
- Gender distribution
- Workload by signing authority
- State-level breakdowns

No direct integration between Python and Power BI is required. Both systems simply read and write to the same database.

---

## Technical Design Decisions

Some engineering decisions made during development:

- **Parallel document generation** using `ProcessPoolExecutor` to render documents across CPU cores simultaneously (v3, Z)
- **In-memory document rendering** using `BytesIO` to avoid writing intermediate files to disk (v3, Z)
- **Merged output generation** with `docxcompose` to combine all documents into a single print-ready file (v3, Z)
- **COM automation of Word Mail Merge** using `win32com` via `DispatchEx` to control Word's native rendering engine from Python (Y, X, O)
- **Batched Mail Merge execution** in groups of 250 records to prevent Word from hanging on large volumes (X, O)
- **Temp Excel bridge** to connect PostgreSQL data to Word's Mail Merge engine, which cannot read from a database directly (O)
- **Batch database updates** with rollback protection to prevent partial status updates on failure
- **Dynamic year injection** to avoid hardcoded template values that break on year transitions
- **Validation before processing** to catch bad data before it enters the document generation stage

---

## Performance

Benchmarked on 108 records (34,020 seconds manual equivalent):

| Version | Time | Efficiency | Rendering Engine | Significance |
|---------|------|------------|------------------|--------------|
| v1 | 50.6s | 99.85% | Python (docxtpl) | Proved automation was already faster than manual work |
| v2 | 9.96s | 99.97% | Python (docxtpl) | Strong raw speed for smaller runs |
| v3 (PLUS) | 28.6s | 99.92% | Python (docxtpl) | Slower than v2, but added validation, merging, and parallel processing |
| Y | 37.85s | 99.89% | Word Mail Merge + Excel | First Python-to-Word Mail Merge crossover |
| X | 27.82s | 99.92% | Word Mail Merge + Excel | Mail Merge branch made scalable with batching |
| Z | 23.06s | 99.93% | Python (docxtpl) + PostgreSQL | Database-backed generation with auto-detection |
| **O** | **16.1s** | **99.95%** | **Word Mail Merge + PostgreSQL** | **Complete system: DB + Mail Merge + write-back + reporting** |

The goal of the later versions was not just raw speed. It was to improve reliability, scalability, and process control while still remaining dramatically faster than manual work. DocuMateO is not the fastest in raw time (v2 is), but it is the only version that closes the full loop: data source, validation, generation, write-back, and reporting in one system.

---

## Ethical Design Note

During development, I explored automating the signing authority's signature on the certificates.

This idea was deliberately dropped for two reasons:

1. A reproduced signature would not look authentic when printed.
2. Automating an officer's identity mark on official documents crosses an ethical boundary without explicit authorisation.

The system therefore only generates the certificate content and leaves signing as a manual step.

---

## Dependencies

```
pandas
python-docx
docxtpl
docxcompose
openpyxl
psycopg2-binary
pywin32
```

`pywin32` is required only for the Mail Merge versions (Y, X, O) and only runs on Windows with Microsoft Word installed.

Install using:

```bash
pip install -r requirements.txt
```

---

## Author

**Angad Kadam**

[LinkedIn](https://linkedin.com/in/angad-kadam-03b606159) | [GitHub](https://github.com/angad08)
