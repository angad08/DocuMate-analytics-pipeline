# Making any kind of document

DocuMate works with any Excel file and any Word template - Birth
Registration, Foreign PCC, a report card. There is nothing to set up per
document. The one settings file, `config.env`, is optional: anything left
blank in it is asked for when DocuMate runs.

[Back to the README](../../README.md)

---

## Running it

```
python main.py v3
```

1. **Pick the Excel file.** A file picker opens. One row = one document.
2. **Pick the Word template.**
3. **Say which column identifies each row** - Passport No, Roll No... -
   from a numbered list. Press Enter to take the one DocuMate suggests.

That's all. It checks the rows, makes the documents into one Word file in
`files/output/`, and marks each row `COMPLETED` in the sheet so it is never
made twice.

Every version (v3, X, Y, Z, O) asks the same way, every run - so
switching from Birth Registration to Foreign PCC is just picking different
files. The pickers open in the folder you used last, and the ID list starts
on last time's answer, so a repeat run is a few clicks and Enter.

The birth register is recognised by its column headings, so it needs no
ID question and gets its own conversions and checks automatically.

---

## Making the Word template

Design the page in Word exactly as it should look, and type the Excel
column name in double curly brackets wherever a value goes:

```
Name:          {{ Applicant_Name }}
Passport No:   {{ Passport_No }}
Issued on:     {{ Issue_Date }}
```

- Spaces and punctuation in a column heading become `_`: the column
  `Passport No` is `{{ Passport_No }}`.
- Names are case-sensitive.
- `{{ year }}` is the current year. `{{ Applicant_Name|upper }}` prints in capitals.
- Type each placeholder in one go - if Word's spell-check splits one, retype it.

This is for v3 and Z. X, Y and O use a **Word Mail Merge** template
instead (Mailings > Select Recipients > Use an Existing List, then Insert
Merge Field) - see [Template setup](TEMPLATES.md).

---

## What gets checked

Before anything is made, DocuMate compares the template with the sheet:

- **A placeholder with no matching column stops the run** and lists both,
  so a typo like `{{ Pasport_No }}` is caught before 200 documents print
  with a gap in them.
- **Only the columns the template prints must be filled in.** Other
  columns in the sheet are ignored.
- **A blank cell in a printed column** is listed, and you're asked whether
  to leave it blank (fine for a Remarks column) or stop and fix it.
- **The ID column** must be filled in and different on every row.
- Dates Excel holds as dates print as `dd/mm/yyyy`; whole numbers print
  without `.0`.

Close the workbook in Excel before running - the `STATUS` column can't be
written to a file that is open.

---

## Running without anyone to ask

Fill the files in `config.env` (in the project folder) and DocuMate uses
them without asking - for the `.exe` files, unattended runs, or just to
skip the pickers:

```
DOCUMATE_EXCEL_PATH=C:\...\FOREIGN_PCC_2026.xlsx
DOCUMATE_TEMPLATE=C:\...\FOREIGN_PCC_FORMAT.docx
DOCUMATE_ID_COLUMN=Passport No
```

Clear a line (`DOCUMATE_EXCEL_PATH=`) and it is asked for again. Every
setting is explained in `config.env` itself.

---

## A database instead of Excel (Z and O)

Z and O read the birth register's database tables. To read a table of your
own, set it in `config.env`:

```
DOCUMATE_DB_TABLE=foreign_pcc
DOCUMATE_DB_ID_COLUMN=passport_no
```

The table needs a text `status` column; DocuMate adds `date_issued`. Every
column is a template field by its name. The template is asked for when the
run starts, as above.
