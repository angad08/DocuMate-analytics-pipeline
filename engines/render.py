"""
The worker that fills one template. Used by template.py.

This has to live in its own small file. When Python starts a worker
process it looks the function up by name and imports the file it's in, so
if this sat next to the Word COM imports, every worker would try to load
Word. Keeping it here keeps the workers light and stops that breaking on
Windows.
"""

import re
from io import BytesIO


def placeholder_name(column):
    """How a column heading is written as a placeholder: "Passport No." -> Passport_No."""
    return re.sub(r"\W+", "_", str(column)).strip("_")


def with_placeholder_names(record):
    """
    The record, plus each field again under a name a placeholder can use.

    A {{ placeholder }} cannot contain a space or a dot, so a column headed
    "Student Name" or "S No." could never be printed. Each such field is
    also offered as Student_Name and S_No, and the original names still
    work for anything that looks them up another way.
    """
    record = dict(record)
    for name in list(record):
        safe = placeholder_name(name)
        if safe and safe != name:
            record.setdefault(safe, record[name])
    return record


def render_record(template_path, record):
    """
    Fill the template with one record and hand back the Word file as bytes.

    Bytes, not a Document object, because Python can't pass Document
    objects between processes. Each worker builds its document in memory -
    no temp files involved.
    """
    from docxtpl import DocxTemplate

    document = DocxTemplate(template_path)
    document.render(with_placeholder_names(record))

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
