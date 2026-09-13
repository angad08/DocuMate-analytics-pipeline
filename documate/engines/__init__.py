"""
How the Word file gets made.

Two files, two engines:

    docxtpl_engine.py    docxtpl - Python fills the template itself  -> v3, Z
    mailmerge_engine.py  Word does it over COM                       -> X, Y, O

Both take the same thing (a list of records) and do the same thing (save
one merged .docx), so the rest of the code doesn't care which one is in use.

The other two files are small helpers both engines share: files.py for
output paths and temp cleanup, word_com.py for the Word constants.
render.py is the worker that fills one template, used by docxtpl_engine.py only.
"""

from documate.engines.mailmerge_engine import MailMergeEngine
from documate.engines.docxtpl_engine import DocxtplEngine
