# Design decisions

The calls that shaped how DocuMate works, and why each one was made.

[Back to the README](../README.md)

---

## Technical Design Decisions

Some engineering decisions made during development:

- **Parallel document generation** using `ProcessPoolExecutor` to render documents across CPU cores simultaneously (v3, Z)
- **In-memory document rendering** using `BytesIO` to avoid writing intermediate files to disk (v3, Z)
- **Merged output generation** with `docxcompose` to combine all documents into a single print-ready file (v3, Z)
- **COM automation of Word Mail Merge** using `win32com` via `DispatchEx` to control Word's native rendering engine from Python (Y, X, O)
- **Batched Mail Merge execution** in groups of 250 records to prevent Word from hanging on large volumes (X, O)
- **Temp CSV bridge** to connect PostgreSQL data to Word's Mail Merge engine, which cannot read from a database directly (O)
- **CSV rather than Excel as the merge source** to stop Word raising its "Select Table" dialog, which blocked every unattended run (Y, X, O)
- **Batch database updates** with rollback protection to prevent partial status updates on failure
- **Dynamic year injection** to avoid hardcoded template values that break on year transitions
- **Validation before processing** to catch bad data before it enters the document generation stage
