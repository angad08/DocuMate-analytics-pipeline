"""
Engines: turn a list of records into one merged Word file.

There are two engines. A version picks one in versions/registry.py.

    docxtpl_engine.py    Python fills the template (docxtpl)      v3, Z
    mailmerge_engine.py  Word fills it with Mail Merge (COM)      X, Y, O

Both have the same method, generate(records, output_path), and both save
one merged .docx, so the pipeline can use either without knowing which.

Helper modules:

    render.py        fills one record into the template (docxtpl only)
    output_paths.py  builds output filenames and cleans up temp files
    word_com.py      named constants for Word's COM options
    pdf.py           saves a PDF copy of the merged .docx (to_pdf=True)
"""

from engines.mailmerge_engine import MailMergeEngine
from engines.docxtpl_engine import DocxtplEngine
