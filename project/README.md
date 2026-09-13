# DocuMate

Document automation and data validation for consular birth registration certificates.

Officers were copying applicant details out of Excel and pasting them into Word
templates, one certificate at a time. I measured that at **315 seconds per
record**. Across the 1,440 records in the working set, that is **126 hours** of
typing that nobody could check and nothing could catch errors in.

DocuMate does the same job in **under two minutes**, and refuses to run at all
if the data going in is bad.

![DocuMate Workflow Architecture](diagram/DocuMate%20Workflow%20Architecture%20Diagram.png)

---

## The problem

Excel was doing two jobs it is not built for: holding the records, and tracking
which ones had been processed. That caused three real problems.

**No validation.** A blank field or a duplicate file number produced a
certificate anyway. Nobody found out until the document was printed and in
someone's hands.

**No visibility.** There was no way to answer "how many are pending" without
opening the sheet and counting.

**No speed.** Every certificate was a person copying nine fields by hand.

The fix had to solve all three. Making the typing faster would have left the
other two alone.

---

## What I built

A pipeline that reads records, checks them, generates the documents, writes the
status back, and leaves the data somewhere Power BI can read it.

The validation runs first, before a single document is created. Five checks:

1. **Dataset exists** before processing starts.
2. **Status column present**, so the filter has something to work on.
3. **Active record filter.** Anything not marked `PRINTED` counts as pending,
   including blanks and half-typed statuses. A record with a missing status is
   never silently skipped.
4. **Mandatory fields**, reporting the exact rows and columns that are empty.
5. **Duplicate file numbers**, caught before they reach a certificate.

If any check fails, it stops and says why. Nothing gets generated.

That last part matters more than the speed. The old process would happily
produce a certificate with a blank parent name. This one will not start.

---

## Results

Benchmarked on 1,440 records against the manual rate of 315 seconds each.

| | Manual | DocuMate (O) |
|---|---:|---:|
| Time | 126 hours | 100 seconds |
| Validation | none | 5 checks, blocking |
| Status tracking | manual | automatic write-back |
| Reporting | none | live Power BI |

Seven versions exist because each one answered a problem the last one had. The
fastest on paper is v1 at 26 seconds, but it does almost nothing: no validation,
no database, no status tracking. O is slower and is still the one I would
deploy, because it closes the loop.

Full benchmark table and the reasoning behind each version:
**[docs/EVOLUTION.md](docs/EVOLUTION.md)**

---

## Quick start

The Excel versions (`v3`, `x`, `y`) need no database and run straight after install.

```bash
pip install -r requirements.txt
python main.py --check     # confirms paths and packages
python main.py --list      # shows the five versions
python main.py v3          # generates from the sample workbook
```

Documents land in `files/output/`. The workbook in
`files/data/` is synthetic test data, so a fresh clone runs with no
setup at all.

The Mail Merge versions (`x`, `y`, `o`) drive Microsoft Word through COM, so
they need Word and Windows. `v3` and `z` render with `docxtpl` and run anywhere
Python does.

### The five versions

| Version | Reads from | Makes the file with | Notes |
|---|---|---|---|
| **v3** | Excel | docxtpl | Deployed at the Consulate in 2025 |
| **x** | Excel | Word Mail Merge | 250 at a time. 1,440 records in about 52s |
| **y** | Excel | Word Mail Merge | All at once. Fine small, hangs large |
| **z** | PostgreSQL | docxtpl | Filters in SQL, polls for new records |
| **o** | PostgreSQL | Word Mail Merge | 100 at a time. The complete version |

```bash
python main.py z --once
python main.py x --poll --interval 300
```

---

## How it is organised

Every version does the same job. They differ in exactly three places: where
records come from, how the Word file is made, and where `PRINTED` gets written
back. So there is one copy of each, and a version just says which combination to
use.

| Folder | What it holds | Edit it when |
|---|---|---|
| `flow/` | The steps every version runs, in order | Changing the sequence or the five checks |
| `sources/` | Where records come from, Excel or PostgreSQL | Adding a data source |
| `engines/` | How the Word file is made, docxtpl or Mail Merge | Adding a rendering method |
| `versions/` | Which combination each version uses | **Adding a version, one entry in `registry.py`** |
| `setup/` | Paths, settings, messages, popups | Changing where things live or what the user sees |
| `tools/` | `seed_database.py`, loads records into PostgreSQL | Seeding the database for `z` and `o` |
| `files/` | What the code reads and writes | Swapping the workbook or the templates |
| `project/` | Docs, diagram, dashboard, schema. Not code | Updating documentation |

There is no code per version. `main.py` reads the key you pass, looks it up in
`versions/registry.py`, plugs the named source and engine into the pipeline in
`flow/pipeline.py`, and runs it.

Adding one is a single entry, no new file:

```python
VERSIONS["p"] = Version(
    key="p", label="P", source="postgres", engine="mailmerge",
    template=TEMPLATE_MAILMERGE, output_prefix="DocuMateP",
    check_records=False, batch_size=500,
    timestamp_format="%d%m%Y_%H%M%S",
)
```

`python main.py p` then works, and the version tests cover it
automatically.

```bash
python -m pytest tests -q     # 39 passed
```

Covers the five checks, status matching, date handling, Serial sorting, the Mail
Merge batch maths at the real 1,440-record volume, reading a database row, and
the version list's own consistency.

---

## Database setup

Only needed for `z` and `o`.

```bash
psql -h YOUR_HOST -U YOUR_USER -d YOUR_DB -f project/database_schema/DocuMate_Data_Schema.sql
```

That creates four normalised tables: `applicant`, `ministryofhomeaffairs`,
`ib_authority`, `state`. Records are pulled with SQL joins across them, and the
pending filter lives in the query's `WHERE` clause, so the database only sends
rows that will actually be used.

Connection settings come from the environment, never from the code:

```bash
cp .env.example .env
```

```
DOCUMATE_DB_HOST=your-host
DOCUMATE_DB_NAME=postgres
DOCUMATE_DB_USER=your-user
DOCUMATE_DB_PASSWORD=your-password
DOCUMATE_DB_PORT=5432
DOCUMATE_DB_SSLMODE=require
```

`.env` is gitignored. If a value is missing, DocuMate stops immediately and
names the variable instead of failing later with a confusing connection error.

Load sample records:

```bash
python -m tools.seed_database --insert
python -m tools.seed_database --select    # what is pending
```

The demo environment uses Supabase, but any PostgreSQL host works.

---

## Power BI

The dashboard reads the same PostgreSQL tables DocuMate writes to. There is no
integration code between them, they just share a database.

Application volume, processing backlog, workload by signing authority, and
state-level breakdowns.

---

## A note on signatures

I looked at automating the signing authority's signature onto the certificates
and decided against it. Two reasons: a reproduced signature does not look right
printed, and putting an officer's identity mark on an official document without
their explicit authorisation is not mine to automate.

DocuMate generates the content. Signing stays manual.

---

## More detail

- **[docs/EVOLUTION.md](docs/EVOLUTION.md)** - all seven versions, the benchmark table, and the two rendering engines
- **[docs/TEMPLATES.md](docs/TEMPLATES.md)** - building the Word templates and the available fields
- **[docs/DESIGN_DECISIONS.md](docs/DESIGN_DECISIONS.md)** - the engineering calls and why
- **[docs/MODULARISATION.md](docs/MODULARISATION.md)** - how the old flat scripts map onto this structure

---

## Dependencies

```
pandas  python-docx  docxtpl  docxcompose  openpyxl  psycopg2-binary  pywin32
```

`pywin32` is only needed for the Mail Merge versions, and only on Windows with
Word installed.

---

## Author

**Angad Kadam**

[LinkedIn](https://linkedin.com/in/angad-kadam-03b606159) | [GitHub](https://github.com/angad08)
