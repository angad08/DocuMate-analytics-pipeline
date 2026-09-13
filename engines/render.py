"""
The worker that fills one template. Used by template.py.

This has to live in its own small file. When Python starts a worker
process it looks the function up by name and imports the file it's in, so
if this sat next to the Word COM imports, every worker would try to load
Word. Keeping it here keeps the workers light and stops that breaking on
Windows.
"""

from io import BytesIO


def render_record(template_path, record):
    """
    Fill the template with one record and hand back the Word file as bytes.

    Bytes, not a Document object, because Python can't pass Document
    objects between processes. Each worker builds its document in memory -
    no temp files involved.
    """
    from docxtpl import DocxTemplate

    document = DocxTemplate(template_path)
    document.render(record)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
