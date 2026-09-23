"""
The docxtpl engine. Used by v3 and Z.

Python fills the Word template itself, with no Word installation needed
(unless to_pdf is on).

Flow
----
1. Each record is sent to render_record() (engines/render.py), which fills
   one copy of the template. Up to max_workers records are filled at the
   same time, each in its own process.
2. The filled documents are collected back in their original order.
3. They are joined into one document with docxcompose, with a page break
   between each record, so every certificate starts on a new page.
4. The merged .docx is saved to output_path.
5. If to_pdf is on, a PDF copy is saved next to it (engines/pdf.py).

A record that fails to fill is reported and skipped; the rest still go
into the file.
"""

import os
from io import BytesIO

from setup import messages
from setup import ui
from engines.output_paths import make_output_folder
from engines.pdf import convert_to_pdf
from engines.render import render_record


class DocxtplEngine:
    """
    Fills the template in parallel, then joins everything into one file.

    template_path   the .docx template with {{ Field_Name }} placeholders
    max_workers     how many records to fill at the same time
    to_pdf          also save a PDF copy of the merged file
    """

    # Name used in progress messages.
    label = "docxtpl engine"

    def __init__(self, template_path, max_workers=10, to_pdf=False):
        self.template_path = template_path
        self.max_workers = max_workers
        self.to_pdf = to_pdf

    def add_page_break(self, document):
        """Add a page break at the end of the document."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        paragraph = document.add_paragraph()
        run = paragraph.add_run()

        # python-docx has no page-break helper for this, so the break is
        # added as raw XML: <w:br w:type="page"/>
        page_break = OxmlElement("w:br")
        page_break.set(qn("w:type"), "page")
        run._r.append(page_break)

    def generate(self, records, output_path):
        """
        Fill every record and save one merged document.

        records       list of dicts, one per certificate
        output_path   where to save the merged .docx

        Returns the 1-based positions of records included in the document.

        Order is kept: results are collected in the order the records were
        submitted, not the order the workers finish, so the output always
        follows the Serial order of the input.
        """
        from concurrent.futures import ProcessPoolExecutor
        from docx import Document
        from docxcompose.composer import Composer

        if not records:
            print(messages.NOTHING_TO_MERGE)
            return []

        total = len(records)
        print(messages.GENERATING.format(count=total))

        merged = None       # the final document; every record is added to it
        composer = None     # docxcompose helper that appends documents
        failed = 0
        successful = []

        with ProcessPoolExecutor(max_workers=self.max_workers) as workers:

            # Step 1: queue every record for filling.
            jobs = []
            for record in records:
                jobs.append(workers.submit(render_record, self.template_path, record))

            # Steps 2 and 3: collect each result in order and append it.
            for number, job in enumerate(jobs, start=1):

                try:
                    finished_bytes = job.result()
                    document = Document(BytesIO(finished_bytes))
                except Exception as error:
                    # Skip this record but keep going with the rest.
                    failed = failed + 1
                    record = records[number - 1]
                    print("\nDocuMate : Error in record " + str(number)
                          + " (Serial: " + str(record.get("Serial"))
                          + ", Name: " + str(record.get("Name")) + ") - " + str(error))
                    continue

                if merged is None:
                    # The first document is the base. Its page setup,
                    # headers and styles are used for the whole file.
                    merged = document
                    composer = Composer(merged)
                else:
                    self.add_page_break(merged)
                    composer.append(document)

                successful.append(number)
                ui.progress(messages.RECORD_PROGRESS.format(done=number, total=total))

        if merged is None:
            print(messages.NOTHING_TO_MERGE)
            return []

        # Step 4: save the merged .docx.
        make_output_folder(os.path.dirname(output_path))
        composer.save(output_path)

        print(messages.SAVED.format(count=total - failed, path=output_path))

        # Step 5: PDF copy, made from the saved .docx.
        if self.to_pdf:
            convert_to_pdf(output_path)

        return successful
