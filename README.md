# DocuMate

Document automation and data validation for consular birth-registration certificates.

Officers were copying applicant details from Excel into Word templates, one certificate at a time. At the measured manual rate of **315 seconds per record**, a 1,440-record workload takes **126 hours** of hands-on work. DocuMate reduces that work to minutes while stopping incomplete or duplicate records before a certificate is produced.

![DocuMate workflow](project/diagram/DocuMate%20Workflow%20Architecture%20Diagram.png)

## The problem

The manual process was slow, error-prone, and difficult to control. Excel was being used both as the data store and as the way to track which records had been processed. That meant there was no dependable validation, no reliable view of backlog, and no consistent way to produce large batches of certificates.

DocuMate automates the full flow:

1. Read applicant records from Excel or PostgreSQL.
2. Validate the records before document generation.
3. Create a single print-ready Word document.
4. Mark successfully processed records as `PRINTED`.
5. Keep the same PostgreSQL data available for reporting.

## Impact

| Measure | Manual process | DocuMate O |
| --- | ---: | ---: |
| 1,440 records | 126 hours | about 100 seconds |
| Validation | manual | five blocking checks |
| Status tracking | manual | write-back to Excel or PostgreSQL |
| Reporting | none | shared PostgreSQL data for Power BI |

The value is not only speed. The validation stage stops a run if required fields are blank, the `STATUS` column is missing, or a file number is duplicated. It prevents a flawed record from becoming a flawed certificate.

## Versions

| Version | Source | Rendering | Purpose |
| --- | --- | --- | --- |
| **v3** | Excel | `docxtpl` | Production Excel workflow with validation and merged output |
| **X** | Excel | Word Mail Merge | Batched at 250 records for large runs |
| **Y** | Excel | Word Mail Merge | One batch; suitable for smaller runs |
| **Z** | PostgreSQL | `docxtpl` | Database-backed workflow with polling |
| **O** | PostgreSQL | Word Mail Merge | Database plus batched Word generation |

The earlier v1 and v2 scripts established the proof of concept. V3 onward is now modular: versions choose a data source and a rendering engine rather than duplicating the full application.

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

The Excel versions (`v3`, `x`, `y`) run from the synthetic workbook in `files/data/`. The Mail Merge versions (`x`, `y`, `o`) require Microsoft Word on Windows. Generated documents go to `files/output/`, which is intentionally ignored by Git.

## Modular structure

```text

├── main.py          command-line runner
├── flow/            validation, pipeline, status updates
├── sources/         Excel and PostgreSQL readers
├── engines/         docxtpl and Word Mail Merge renderers
├── versions/        v3, X, Y, Z and O settings
├── setup/           paths, configuration, messages, UI
├── files/           synthetic data, templates, generated output
└── project/         documentation, schema and workflow diagram
```

To change a version, edit [the version registry](versions/registry.py). To change the workflow shared by every version, edit [the pipeline](flow/pipeline.py). Connection and file-path settings belong in [configuration](setup/config.py), while private values belong in `.env`.

## Database setup

Z and O use PostgreSQL. Create the schema with:

```bash
psql -h YOUR_HOST -U YOUR_USER -d YOUR_DB -f project/database_schema/DocuMate_Data_Schema.sql
```

Copy `.env.example` to `.env` and enter the database connection values. `.env` is ignored by Git; never place passwords or connection credentials in source files.

```text
DOCUMATE_DB_HOST=your-host
DOCUMATE_DB_NAME=postgres
DOCUMATE_DB_USER=your-user
DOCUMATE_DB_PASSWORD=your-password
DOCUMATE_DB_PORT=5432
DOCUMATE_DB_SSLMODE=require
```

## Further reading

- [Evolution and benchmarks](project/docs/EVOLUTION.md)
- [Template setup](project/docs/TEMPLATES.md)
- [Engineering decisions](project/docs/DESIGN_DECISIONS.md)
- [Modularisation map](project/docs/MODULARISATION.md)

## Author

**Angad Kadam**

[LinkedIn](https://linkedin.com/in/angad-kadam-03b606159) · [GitHub](https://github.com/angad08)
