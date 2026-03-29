"""
====================================================================
 DocuMate v1.0 — Birth Registration
====================================================================

 This is where it all started.

 Not a product, not a platform — just a script born out of
 frustration with repetitive, Excel-heavy, manual birth
 registration work. A first attempt at saying:
 "There has to be a better way."

 Everything that came after traces back to this file.

--------------------------------------------------------------------
 WHAT IT DOES
--------------------------------------------------------------------

 1. READS DATA
    Source: "DocuMate_DataFrame.xlsx" → DocuMateSRC

 2. FILTERS ACTIVE CASES
    Only rows where STATUS == "IN PROCESS"
    (Note: filtering wasn't in the very first draft —
     it was the first improvement once the idea proved useful.)

 3. FORMATS DATES
    Converts date columns to dd/mm/yyyy so documents
    look official, not like raw Excel exports.

 4. GENERATES WORD DOCUMENTS
    Template:  "DOCUMENT_TEMPLATE_FILE.docx"
    Output:    output_files/
    Naming:    Serial "123/2025" → "123.docx"
    One document per record.

 5. PRINTS SUMMARY
    Total records processed + time taken.

--------------------------------------------------------------------
 LIMITATIONS (v1 — intentionally minimal)
--------------------------------------------------------------------

 - No error handling, no validations, no defensive checks
 - No user interface
 - Purpose was simple: prove the idea works, reduce manual
   effort, turn Excel rows into real documents automatically

 Later versions add safety, structure, logging, and scale.
 This file is the foundation.

====================================================================
"""


import os

import pandas as pd
from docxtpl import DocxTemplate
import time

start_time = time.time()

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Load Excel file and specify sheet name
df = pd.read_excel(os.path.join(base_dir, "data", "DocuMate_DataFrame.xlsx"), sheet_name="DocuMateSRC")

# ✅ Filter only rows where Status = "IN PROCESS"
df = df[df["STATUS"].str.upper() == "IN PROCESS"]

# Convert any datetime columns to string format dd/mm/yyyy
date_cols = ["Registration_date"]  # add more if needed
for col in date_cols:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y")

for idx, row in df.iterrows():
    doc = DocxTemplate(os.path.join(base_dir, "templates", "DOCUMENT_TEMPLATE_FILE.docx"))

    # Create context from Excel row
    context = row.to_dict()
    # Render the template
    doc.render(context)

    # Use only the number before "/" from Serial as filename
    serial = str(row.get("Serial", f"row{idx+1}")).split("/")[0]

    output_filename = os.path.join(base_dir, "output_files", f"{serial}.docx")
    doc.save(output_filename)
    print(f"✅ Created: {serial}.docx")

# ✅ Print summary
print(f"Total records processed: {len(df)}")
print(f"Time taken: {time.time() - start_time:.2f} seconds")