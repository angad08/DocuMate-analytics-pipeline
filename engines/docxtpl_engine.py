"""
The docxtpl engine. Shared by v3 and Z.

Python fills the template itself, several records at a time in separate
processes, then joins the results together with a page break between each.

v3 and Z each had their own copy of this. The only differences were the
docstrings and the default filename. Once both sources started handing over
the same list of records, there was no reason for two copies.
"""

import os
from io import BytesIO

from setup import messages
from setup import ui
from engines.output_paths import make_output_folder
from engines.render import render_record


class DocxtplEngine:
    """Fills the template in parallel, then joins everything into one file."""

    # Shown in messages.
    label = "docxtpl engine"

    def __init__(self, template_path, max_workers=10):
        self.template_path = template_path
        self.max_workers = max_workers

    def add_page_break(self, document):
        """Add a page break, so each certificate starts on a fresh page."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        paragraph = document.add_paragraph()
        run = paragraph.add_run()

        page_break = OxmlElement("w:br")
        page_break.set(qn("w:type"), "page")
        run._r.append(page_break)

    def generate(self, records, output_path):
        """
        Fill every record and save one merged document.

        About the ordering: we hand all the records to the workers, then
        collect the results in the order we submitted them, NOT the order
        they finish. So the document comes out in Serial order no matter
        which worker happens to finish first.
        """
        from concurrent.futures import ProcessPoolExecutor
        from docx import Document
        from docxcompose.composer import Composer

        if not records:
            print(messages.NOTHING_TO_MERGE)
            return output_path

        total = len(records)
        print(messages.GENERATING.format(count=total))

        merged = None       # the document everything gets added to
        composer = None     # the tool that does the adding
        failed = 0

        with ProcessPoolExecutor(max_workers=self.max_workers) as workers:

            jobs = []
            for record in records:
                jobs.append(workers.submit(render_record, self.template_path, record))

            for number, job in enumerate(jobs, start=1):

                try:
                    finished_bytes = job.result()
                    document = Document(BytesIO(finished_bytes))
                except Exception as error:
                    # One bad record shouldn't cost us the other 299.
                    failed = failed + 1
                    print("\nDocuMate : Error in record " + str(number) + " - " + str(error))
                    continue

                if merged is None:
                    # The first document becomes the base, so its page setup,
                    # headers and styles carry through to the final file.
                    merged = document
                    composer = Composer(merged)
                else:
                    self.add_page_break(merged)
                    composer.append(document)

                ui.progress(messages.RECORD_PROGRESS.format(done=number, total=total))

        if merged is None:
            print(messages.NOTHING_TO_MERGE)
            return output_path

        make_output_folder(os.path.dirname(output_path))
        composer.save(output_path)

        print(messages.SAVED.format(count=total - failed, path=output_path))
        return output_path
