# DocuMate

DocuMate is a document automation and data validation system I built to replace a slow, manual workflow that relied heavily on Excel copy-paste operations.

The project started as a small Python script to reduce repetitive work and eventually evolved into a structured system that combines **data validation, document automation, PostgreSQL data storage, and Power BI reporting**.

What began as a proof-of-concept became a platform that can process thousands of records reliably while maintaining data quality checks and operational visibility.

Manual processing of roughly **1,440 records typically required nearly two months of manual effort**.
DocuMate reduces that to **a few minutes of automated processing**, with safeguards that prevent bad data from entering the document generation stage.

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

---

## What I Built

DocuMate gradually evolved into a system that automates the entire pipeline:

1. Reads applicant records from **Excel or PostgreSQL**
2. Runs a **multi-stage data validation process**
3. Maps validated data into **Word document templates**
4. Generates certificates automatically
5. Merges all documents into a **single print-ready file**
6. Updates processing status in Excel or the database
7. Exposes the same data to **Power BI dashboards for reporting**

Instead of manually generating hundreds or thousands of files, the system processes the records in bulk while enforcing validation rules.

---

## How the System Evolved

The project went through several iterations as new problems became visible.

| Version | Purpose | Key Improvements |
|---------|---------|-----------------|
| **v1.0** | Proof of concept | Basic script to read Excel and generate Word documents |
| **v2.0** | Stabilisation | Introduced OOP structure, error handling, and portable file paths |
| **v3.0 (PLUS)** | Data reliability | Added 5-stage validation, merged document output, parallel processing |
| **v4.0 / X** | System architecture | Migrated to PostgreSQL, SQL joins, batch updates, auto-detection mode |

Each version addressed limitations discovered in the previous one rather than starting from scratch.

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
│   ├── DocuMateX.py
│   └── loadData.py
│
├── sql/
│   └── DocuMate_Data_Schema.sql
│
├── templates/
│   └── DOCUMENT_TEMPLATE_FILE.docx
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

1. **Dataset existence check** — Confirms that data is available before processing begins.
2. **Status column verification** — Ensures required fields exist in the dataset.
3. **Active record check** — Filters only records marked "IN PROCESS".
4. **Mandatory field validation** — Identifies rows with missing values and reports the exact fields.
5. **Duplicate detection** — Prevents duplicate file numbers from entering the system.

If any validation step fails, processing stops immediately and the issue is reported.

This prevents bad data from generating incorrect certificates.

---

## PostgreSQL Architecture

The final version of DocuMate migrated the data layer to PostgreSQL to move away from Excel acting as a database.

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

The demonstration environment uses Supabase (managed PostgreSQL), but DocuMate works with any PostgreSQL-compatible host.

---

## Database Setup

Create a PostgreSQL database and run the included schema file:

```bash
psql -h YOUR_HOST -U YOUR_USER -d YOUR_DB -f sql/DocuMate_Data_Schema.sql
```

Update the database credentials in:

```
src/DocuMateX.py
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

### Excel versions (v1, v2, v3)

```bash
python src/DocuMate_v1_0.py
python src/DocuMate_v2_0.py
python src/DocuMatePLUS_v3_0.py
```

These versions read records from `data/DocuMate_DataFrame.xlsx`.

### Database version (v4 / X)

```bash
python src/DocuMateX.py
```

DocuMateX can run in **auto-detection mode**, polling the database for "IN PROCESS" records every 30 seconds. This allows new records to be processed automatically.

---

## Power BI Integration

The Power BI dashboard connects to the same PostgreSQL database used by DocuMate.

DocuMate performs the processing and updates record status. Power BI reads the same tables to generate operational dashboards.

Typical insights include:

- Application volume trends
- Processing backlog
- Gender distribution
- Workload by signing authority
- State-level breakdowns

No direct integration between Python and Power BI is required. Both systems simply read and write to the same database.

---

## Technical Design Decisions

Some engineering decisions made during development:

- **Parallel document generation** using `ProcessPoolExecutor`
- **In-memory document rendering** using `BytesIO`
- **Merged output generation** with `docxcompose`
- **Batch database updates** with rollback protection
- **Dynamic year injection** to avoid hardcoded template values
- **Validation before processing** to prevent incorrect output

These design choices prioritise reliability and predictable processing over raw speed.

---

## Performance

| Version | Records | Time | Output |
|---------|---------|------|--------|
| v1 | 1,440 | ~32 seconds | Individual files |
| v2 | 1,440 | ~31 seconds | Individual files |
| v3 | 1,440 | ~211 seconds | Single merged document |
| v4/X | 1,440 | ~213 seconds | Single merged document |

Later versions take longer because all documents are merged into a single output file, which grows progressively larger as records are added.

Even with this overhead, the automated workflow is dramatically faster than manual processing.

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
```

Install using:

```bash
pip install -r requirements.txt
```

---

## Author

**Angad Kadam**

[LinkedIn](https://linkedin.com/in/angad-kadam-03b606159) | [GitHub](https://github.com/angad08)
