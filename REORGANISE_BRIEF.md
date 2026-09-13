# Brief: reorganise the documate/ folder

Hand this file to Claude Code. Everything below is decided — please follow it
rather than proposing a different structure.

**Hard rule: this is a move-and-rename job only. No behaviour changes.**
The 39 tests must still pass and `--check` must still work when you're done.

---

## Where it is

```
C:\Users\Angat\OneDrive\Documents\DocuMate-analytics-pipeline-main\DocuMate-analytics-pipeline-main\
```

Note the folder is nested inside itself — that's from a zip extraction. Moving
the inner folder up a level and deleting the outer shell would be welcome, but
do that **last**, after everything else works.

---

## Target layout

```
documate/
├── main.py                  the only file you run
│
├── flow/                    THE STEPS EVERY VERSION RUNS
│   ├── __init__.py
│   ├── pipeline.py            the order they happen in
│   ├── checks.py              the five checks   (renamed from validation.py)
│   └── update_status.py       ask, then mark printed
│
├── sources/                 WHERE RECORDS COME FROM
│   ├── __init__.py
│   ├── excel.py               v3, X, Y
│   ├── postgres.py            Z, O
│   └── queries.py
│
├── engines/                 HOW THE WORD FILE IS MADE
│   ├── __init__.py
│   ├── docxtpl_engine.py      v3, Z
│   ├── mailmerge_engine.py    X, Y, O
│   ├── render.py
│   ├── word_com.py
│   └── output_paths.py        (renamed from files.py)
│
├── versions/                WHICH COMBINATION EACH VERSION USES
│   ├── __init__.py            turns the names into objects
│   └── registry.py            ← the table you edit to add a version
│
├── setup/                   settings, messages, popups
│   ├── __init__.py
│   ├── config.py
│   ├── messages.py
│   └── ui.py
│
├── tools/
│   ├── __init__.py
│   └── seed_database.py
│
├── tests/
│   └── test_documate.py
│
├── files/                   what the code reads and writes
│   ├── data/                  DocuMate_DataFrame.xlsx, test_data/
│   ├── templates/             the two .docx templates
│   └── output/                generated documents   (renamed from output_files)
│
└── project/                 not code
    ├── README.md
    ├── docs/
    ├── diagram/
    ├── dashboard/
    └── database_schema/
```

`__init__.py` is needed in `flow/` and `setup/` — they're new packages.

---

## The moves

| From | To |
|---|---|
| `documate/pipeline.py` | `documate/flow/pipeline.py` |
| `documate/validation.py` | `documate/flow/checks.py` |
| `documate/update_status.py` | `documate/flow/update_status.py` |
| `documate/config.py` | `documate/setup/config.py` |
| `documate/messages.py` | `documate/setup/messages.py` |
| `documate/ui.py` | `documate/setup/ui.py` |
| `documate/engines/files.py` | `documate/engines/output_paths.py` |
| `documate/data/` | `documate/files/data/` |
| `documate/templates/` | `documate/files/templates/` |
| `documate/output_files/` | `documate/files/output/` |
| `documate/README.md` | `documate/project/README.md` |
| `documate/docs/` | `documate/project/docs/` |
| `documate/diagram/` | `documate/project/diagram/` |
| `documate/dashboard/` | `documate/project/dashboard/` |
| `documate/database_schema/` | `documate/project/database_schema/` |

Stay where they are: `main.py`, `__init__.py`, `sources/`, `engines/`,
`versions/`, `tools/`, `tests/`.

### Two renames inside the code, not just the filename

- `validation.py` → `checks.py`. The module is imported as
  `from documate import validation` and called as `validation.check_records(...)`
  in `pipeline.py`, and imported directly in the tests. Update both.
- `engines/files.py` → `engines/output_paths.py`. It exports
  `make_output_folder`, `build_output_name`, `delete_temp_file`.

Do **not** rename the functions themselves. The names were settled with the
user: `load_records`, `transform_records`, `validate_template_file`,
`update_status`, `check_records`, `plan_batches`, `build_record`,
`mark_printed`, `build_output_name`.

---

## Delete these

```
documate/__pycache__/          and every other __pycache__ anywhere
documate/engines/files.py      after the rename
CLEANUP_OLD_FILES.bat          at the root, if still there
main.py                        at the ROOT, if still there (the live one is documate/main.py)
```

---

## Code changes the moves force

### 1. `setup/config.py` — path functions

`project_root()` walks up from `__file__`. It is now one level deeper
(`documate/setup/config.py` instead of `documate/config.py`), so it needs
**one more** `os.path.dirname(...)`.

`package_folder()` currently returns `project_root()/documate`. The three
asset folders moved into `files/`, so:

```python
def files_folder():
    return os.path.join(package_folder(), "files")

def data_folder():
    return os.path.join(files_folder(), "data")

def templates_folder():
    return os.path.join(files_folder(), "templates")

def output_folder():
    return os.path.join(files_folder(), "output")
```

Verify by running `python documate\main.py --check` — it prints every path
and says whether each exists. That's the fastest way to confirm this bit.

### 2. Imports

Mechanical. Roughly:

```
from documate import config            →  from documate.setup import config
from documate import messages          →  from documate.setup import messages
from documate import ui                →  from documate.setup import ui
from documate import validation        →  from documate.flow import checks
from documate.validation import CheckFailed
                                       →  from documate.flow.checks import CheckFailed
from documate.pipeline import Pipeline →  from documate.flow.pipeline import Pipeline
from documate.update_status import update_status
                                       →  from documate.flow.update_status import update_status
from documate.engines.files import ... →  from documate.engines.output_paths import ...
```

Call sites using `validation.check_records(...)`, `validation.format_dates(...)`,
`validation.sort_by_serial(...)`, `validation.to_records(...)` become
`checks.…`.

### 3. `.env` location

`config.env_file()` looks in `project_root()` — the folder that **contains**
`documate/`, not inside it. There is currently **no `.env` anywhere**, so
versions `z` and `o` will fail. Create one from `documate/.env.example` and
put it at the project root. Also move `.env.example` and `.gitignore` back out
to the root, next to it — they belong beside the file they describe.

The six variables are `DOCUMATE_DB_HOST`, `DOCUMATE_DB_NAME`,
`DOCUMATE_DB_USER`, `DOCUMATE_DB_PASSWORD`, `DOCUMATE_DB_PORT`,
`DOCUMATE_DB_SSLMODE`. Do not print the password anywhere.

### 4. `tests/test_documate.py`

It already walks up to find the folder containing `documate/`, so it does not
need path changes. Only its imports change (see above).

### 5. Docs

Update paths in `project/README.md` and `project/docs/MODULARISATION.md`:
the folder tree, the architecture image path, the `psql -f ...schema.sql`
path, and the run/test commands.

---

## When you're done

```
python -m pytest documate\tests -q      → 39 passed
python documate\main.py --check         → every path "ok", no MISSING
python documate\main.py --list          → v3, x, y, z, o
python documate\main.py v3              → answer "no" at the prompt
```

`v3` should find **121 pending records** out of 1,440 and write a document
into `documate\files\output\`. If the count differs, something in the move
broke the filtering — stop and say so rather than adjusting the tests.

---

## Context worth having

- Five versions, one pipeline. They differ in exactly three places: the data
  source, the document engine, and where `PRINTED` is written back. That's why
  there are two files in `sources/` and two engines — not five of anything.
- A version is one entry in `versions/registry.py`. There is no code per
  version.
- `render.py` must stay its own module. A multiprocessing worker is imported
  by name in the child process, so if it sat next to the Word COM imports
  every worker would try to load Word.
- `engines/mailmerge_engine.py` has a long comment at the top explaining why
  Word is handed a `.csv` and not the `.xlsx`. That was expensive to work out.
  Keep it.
- `sources/excel.py` has a note on `mark_printed()` saying it updates every
  non-printed row in the sheet, not just this run's. That is deliberate,
  matches the original behaviour, and is a known trade-off. Don't "fix" it.
- All five versions ran successfully on this machine today. If one stops
  working after the move, the move is the cause.
