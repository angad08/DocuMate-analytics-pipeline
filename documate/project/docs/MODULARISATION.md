# How the old scripts map onto the new ones

The project used to be five scripts in `src/`, 3,396 lines, each one a full
copy of the whole pipeline. `documate/` is the same behaviour with the copies
removed.

---

## Why it collapsed the way it did

Put the five side by side and they differ along exactly two lines:

|                      | reads Excel | reads the database |
|----------------------|-------------|--------------------|
| **docxtpl engine**   | v3          | Z                  |
| **Word Mail Merge**  | X, Y        | O                  |

Everything else was the same code, pasted. Counted:

- the five checks appeared **three times**, identical
- the Excel status update appeared **three times**, identical
- the database read and update appeared **twice**, identical apart from
  comment punctuation
- the docxtpl engine appeared **twice** — the difference was docstrings and
  the default filename
- the Mail Merge engine appeared **three times** — the only real difference
  was the batch size: 250, none, 100
- `run()` appeared **four times** — the difference was the print statements

So the things that actually differ come to: two sources, two engines, three
numbers.

---

## Where everything went

| Old code (in every one of the five files) | New home |
|---|---|
| `load_data()` reading Excel | `sources/excel.py` |
| `load_data()` reading the database | `sources/postgres.py` + `sources/queries.py` |
| turning a database row into template fields | `sources/postgres.py`, `build_record()` |
| `filter_records()` — the five checks | `flow/checks.py`, `check_records()` |
| `format_dates()` | `flow/checks.py` |
| the Serial sorting block | `flow/checks.py`, `sort_by_serial()` |
| `add_page_break()` | `engines/docxtpl_engine.py` |
| `process_record()` | `engines/render.py` |
| `generate_and_merge_documents()` — docxtpl | `engines/docxtpl_engine.py` |
| `generate_and_merge_documents()` — Mail Merge | `engines/mailmerge_engine.py` |
| working out the batch ranges | `engines/mailmerge_engine.py`, `plan_batches()` |
| `update_excel_STATUS()` | `sources/excel.py`, `mark_printed()` |
| `update_sql_status()` | `sources/postgres.py`, `mark_printed()` |
| `validate_environment()` | `flow/pipeline.py`, `validate_template_file()` |
| `run()` | `flow/pipeline.py`, `run()` |
| the yes/no prompt + status write-back | `flow/update_status.py` |
| `start()` — the polling loop | `flow/pipeline.py`, `start()` |
| `MessageBoxW` calls (about 30 of them) | `setup/ui.py`, `notify()` |
| the settings at the bottom of each file | `versions/registry.py` |
| `loadData.py` | `tools/seed_database.py` |

`v1` and `v2` have no new home. They were replaced by v3 and aren't part of
the five-version set. Their story stays in the README chronology.

---

## Adding a version

One entry in `versions/registry.py`. Say you want database + Mail Merge in
batches of 500:

```python
VERSIONS["p"] = Version(
    key="p",
    label="P",
    source="postgres",
    engine="mailmerge",
    template=TEMPLATE_MAILMERGE,
    output_prefix="DocuMateP",
    check_records=False,
    batch_size=500,
    timestamp_format="%d%m%Y_%H%M%S",
)
```

`python main.py p` works straight away, and the version tests cover it.

---

## What actually changed in behaviour

Nine things. The first four are bugs that were in the originals.

1. **Output filenames used to freeze.** The old code put the timestamp in a
   *default argument*:

   ```python
   def generate_and_merge_documents(self, merged_filename=f"...{datetime.now()...}.docx"):
   ```

   Python works out a default argument **once**, when the file is first
   imported — not each time you call it. Z and O check for new records every
   30 seconds, so every run reused the timestamp from startup and **overwrote
   the previous document**. Worth knowing about even if you never use the new
   code.

2. **Unreadable dates printed the word "NaT".** `pd.to_datetime(errors="coerce")`
   gives back NaT, and `.strftime` turned that into the literal text `NaT` in
   the certificate. It comes out empty now.

3. **The pre-flight check now runs for every version.** Only Z and O checked
   the template existed. A wrong path in v3, X or Y only showed up after the
   sheet had been read and Word was already open.

4. **The `year` field now reaches every version.** v3 added it so templates
   would stop hardcoding the year — that hardcoding is what caused the missed
   update at the 2025→2026 rollover. X, Y, Z and O never got the fix. Turn it
   off with `INJECT_YEAR = False` in `setup/config.py`.

5. **One set of messages.** The wording differences between versions carried
   no information. All in `setup/messages.py` now.

6. **Database settings moved to `.env`.** `DocuMateZ.py` and `DocuMateO.py`
   had the Supabase host, database and user written into a git-tracked file.
   Only the password was read from the environment. All of it comes from
   `.env` now, and a missing value stops the run and names the variable.

7. **Y writes one temp file.** Y used to save Word's output straight to the
   final path. It now goes through the same single-batch path as X and O,
   which writes one temp file and moves it. Same result, one less code path.

8. **Records convert with `to_dict("records")`, not `itertuples()`.**
   `itertuples` quietly renames any column that isn't a valid Python name to
   `_1`, `_2` and so on — so renaming a column to have a space in it would
   have blanked that field with no error. Checked as identical on the current
   sheet.

9. **A failed record no longer throws off the count.** The docxtpl engine
   already caught per-record errors; the final number now reflects what was
   actually produced.

**Deliberately left alone:** the Excel status update still marks *every*
non-printed row in the sheet, not just the rows from this run. That is how it
always worked, and changing it is a real change rather than a tidy-up. It
matters if someone adds a row while documents are being made — that row gets
marked printed without a certificate. The database version doesn't have this
problem because it updates by Serial number. Noted in `sources/excel.py` right
at the method.

---

## What was checked

```
python -m pytest documate/tests -q     # 39 passed
```

Beyond that, checked against the real `documate/files/data/DocuMate_DataFrame.xlsx`
(1,440 rows, 121 pending):

- **Same records.** The original v3 filtering, date formatting and sorting was
  run verbatim next to the new code. Same 121 records, same order, **zero
  differences in any field.**
- **End to end.** `python main.py v3` produced the 121-certificate document in
  5 seconds, correct page breaks.
- **Status write-back.** On a copy of the workbook: 121 pending → 0,
  `Date_Printed` stamped, all three sheets and their formulas intact.

**Not checked, and can't be from a Linux machine:** anything involving Word
(X, Y, O) or a live database (Z). The COM code is carried over line for line,
but run X and Z on the Consulate machine before relying on them.
