"""
The Word Mail Merge engine. Used by X, Y and O.

Word does the filling, using its built-in Mail Merge, driven from Python
over COM. Needs Windows, Microsoft Word and pywin32.

Flow
----
1. The records are written to a temporary .csv in the output folder
   (write_merge_csv). This is the data source Word merges from.
2. A separate copy of Word is started and the Mail Merge template (the
   _MM.docx file) is opened read-only.
3. The template is connected to the .csv.
4. The records are merged in batches of batch_size (plan_batches). Each
   batch is saved as a temporary .docx.
5. The batch files are joined into one document, with a section break
   between batches, and saved to output_path.
6. If to_pdf is on, a PDF copy is exported while Word is still open.
7. Word is closed and every temp file (the .csv and the batch files) is
   deleted, whether the run worked or not.

Batching
--------
Word can stop responding when asked to build one very large document in a
single merge (it hung on 1,440 records in testing). Merging a few hundred
at a time and joining the results avoids that. batch_size=None merges
everything in one go, which is what version Y does.

Why a .csv and not the Excel workbook
-------------------------------------
When Word merges directly from an .xlsx, it opens the workbook through the
Excel OLEDB driver, which treats every sheet as a table. Word then shows a
"Select Table" dialog on every run and waits for someone to click it, which
blocks an unattended run. Settings such as SQLStatement, SubType or a
single-sheet workbook do not stop the dialog.

A .csv holds exactly one table, so Word connects straight away with no
dialog. Keep OpenDataSource to just the .csv filename: adding a Connection
string or SQLStatement brings the dialog back.

The .csv only contains the records for this run, so the merge always covers
records 1 to N with no gaps.
"""

import os
import shutil
import time

import pandas as pd

from setup import messages
from engines import word_com
from engines.output_paths import delete_temp_file
from engines.output_paths import make_output_folder
from engines.pdf import export_document


# Records per Word merge. See "Batching" above.
DEFAULT_BATCH_SIZE = 250

# Pauses (in seconds) that give Word time to finish one COM call before the
# next. Without them the merge can fail intermittently on slower machines.
SETTLE = 0.5
STARTUP = 1.0

# The temporary merge source, written to the output folder and deleted
# after every run.
TEMP_CSV_NAME = "__temp_merge_source.csv"


def plan_batches(total, batch_size):
    """
    Split records 1..total into (first, last) ranges for Word to merge.

        plan_batches(600, 250)   ->  [(1, 250), (251, 500), (501, 600)]
        plan_batches(600, None)  ->  [(1, 600)]
        plan_batches(0, 250)     ->  []

    Both ends of each range are included. If batch_size is None or larger
    than total, one range covers everything.
    """
    if total <= 0:
        return []

    if not batch_size or batch_size >= total:
        return [(1, total)]

    batches = []
    for start in range(1, total + 1, batch_size):
        end = min(start + batch_size - 1, total)
        batches.append((start, end))

    return batches


def write_merge_csv(records, folder):
    """
    Write the records to the temporary .csv that Word merges from.

    Returns the path of the .csv.

    The file is saved as utf-8-sig (UTF-8 with a byte order mark). Word
    needs the mark to read the header row and any non-English characters
    correctly.
    """
    path = os.path.join(os.path.abspath(folder), TEMP_CSV_NAME)
    pd.DataFrame(list(records)).to_csv(path, index=False, encoding="utf-8-sig")
    return path


class MailMergeEngine:
    """
    Runs Word's Mail Merge through COM and saves one merged document.

    template_path   the Mail Merge template (the _MM.docx file)
    batch_size      records per merge; None merges all at once
    visible         show the Word window while it works
    to_pdf          also save a PDF copy of the merged file
    """

    # Name used in progress messages.
    label = "Word Mail Merge"

    def __init__(self, template_path, batch_size=DEFAULT_BATCH_SIZE, visible=True, to_pdf=False):
        self.template_path = os.path.abspath(template_path)
        self.batch_size = batch_size
        self.visible = visible
        self.to_pdf = to_pdf

    def generate(self, records, output_path):
        """
        Merge every record and save one document. The steps are listed at
        the top of this file.

        records       list of dicts, one per certificate
        output_path   where to save the merged .docx

        Returns output_path.
        """

        if not records:
            print(messages.NOTHING_TO_MERGE)
            return output_path

        pythoncom, win32com = word_com.load_word()

        total = len(records)
        output_path = os.path.abspath(output_path)
        folder = make_output_folder(os.path.dirname(output_path))

        if not os.path.exists(self.template_path):
            raise FileNotFoundError(
                messages.TEMPLATE_MISSING.format(path=self.template_path)
            )

        # Step 1: the data source.
        csv_path = write_merge_csv(records, folder)
        print(messages.MM_SOURCE_BUILT.format(count=total))

        batches = plan_batches(total, self.batch_size)
        print(messages.GENERATING.format(count=total))

        word = None
        template = None
        combined = None
        batch_files = []

        try:
            pythoncom.CoInitialize()

            print(messages.MM_STARTING)

            # Step 2: start Word and open the template.
            # DispatchEx starts a new, separate Word process. Dispatch would
            # attach to a Word window the user already has open, and the
            # merge could then pick up that window's state.
            word = win32com.DispatchEx("Word.Application")
            word.Visible = self.visible
            word.DisplayAlerts = 0
            word.AutomationSecurity = word_com.DISABLE_MACROS
            time.sleep(STARTUP)

            template = word.Documents.Open(
                self.template_path,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            time.sleep(SETTLE)

            print(messages.MM_CONNECTING)

            # Step 3: connect the template to the .csv. Only the filename is
            # given, with no Connection or SQLStatement (see "Why a .csv" at
            # the top of this file).
            template.MailMerge.OpenDataSource(
                Name=csv_path,
                ConfirmConversions=False,
                ReadOnly=True,
                LinkToSource=True,
                AddToRecentFiles=False,
                Revert=False,
            )

            print(messages.MM_CONNECTED)
            time.sleep(SETTLE)

            template.MailMerge.MainDocumentType = word_com.FORM_LETTERS
            template.MailMerge.Destination = word_com.SEND_TO_NEW_DOCUMENT
            template.MailMerge.SuppressBlankLines = True

            # Step 4: merge each batch into its own temp file.
            for number, (first, last) in enumerate(batches, start=1):

                if len(batches) > 1:
                    print(messages.MM_BATCH.format(
                        number=number,
                        total=len(batches),
                        first=first,
                        last=last,
                    ))
                else:
                    print(messages.MM_RANGE.format(first=first, last=last))

                batch_path = self.run_one_batch(word, template, first, last, folder, number)
                batch_files.append(batch_path)

            # The template is no longer needed. Close it before joining so
            # Word is not holding it open.
            self.close_document(template)
            template = None

            # Step 5: join the batches into the final file.
            combined = self.join_batches(word, batch_files, output_path)

            print(messages.SAVED.format(count=total, path=output_path))

            # Step 6: PDF copy, using the Word that is already running.
            # combined is None when there was only one batch, because that
            # file was moved into place rather than opened, so open it here.
            if self.to_pdf:
                try:
                    if combined is None:
                        combined = word.Documents.Open(
                            output_path,
                            ReadOnly=True,
                            AddToRecentFiles=False,
                        )
                    export_document(combined, output_path)
                except Exception as error:
                    print(messages.PDF_FAILED.format(error=error))

            return output_path

        except Exception as error:
            print(messages.MM_ERROR.format(error=error))
            raise

        finally:
            # Step 7: close Word and clean up, even after an error.
            self.close_document(combined)
            self.close_document(template)
            self.quit_word(word)

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            # The .csv and batch files contain personal data, so they are
            # always deleted.
            for path in batch_files + [csv_path]:
                delete_temp_file(path)

    def run_one_batch(self, word, template, first, last, folder, number):
        """
        Merge records first..last and save the result as a temp .docx.

        Returns the path of the temp file (__temp_batch_<number>.docx).
        """

        template.MailMerge.DataSource.FirstRecord = int(first)
        template.MailMerge.DataSource.LastRecord = int(last)
        template.MailMerge.Execute(Pause=False)
        time.sleep(SETTLE)

        # Execute() opens the merge result as a new document, which becomes
        # Word's active document.
        result = word.ActiveDocument

        path = os.path.join(os.path.abspath(folder), "__temp_batch_" + str(number) + ".docx")
        result.SaveAs2(path, FileFormat=word_com.FORMAT_DOCX)

        self.close_document(result)
        return path

    def join_batches(self, word, batch_files, output_path):
        """
        Join the batch files into one document saved at output_path.

        Returns the open combined document, or None when there was only one
        batch. A single batch file is simply moved to output_path, which
        avoids opening and re-saving it for no reason.
        """
        if len(batch_files) == 1:
            shutil.move(batch_files[0], output_path)
            del batch_files[:]
            return None

        combined = word.Documents.Open(batch_files[0])

        # Append each remaining batch after a next-page section break.
        # Content.End - 1 is the position just before the final paragraph
        # mark, which is the end of the document's text.
        for path in batch_files[1:]:
            end = combined.Content.End - 1
            combined.Range(end, end).InsertBreak(word_com.SECTION_BREAK_NEXT_PAGE)

            end = combined.Content.End - 1
            combined.Range(end, end).InsertFile(path)

        combined.SaveAs2(output_path, FileFormat=word_com.FORMAT_DOCX)
        return combined

    def close_document(self, document):
        """Close a document without saving. Does nothing if it is None or already closed."""
        if document is None:
            return
        try:
            document.Close(word_com.DO_NOT_SAVE_CHANGES)
        except Exception:
            pass

    def quit_word(self, word):
        """Quit Word. Does nothing if it is None or already closed."""
        if word is None:
            return
        try:
            word.Quit()
        except Exception:
            pass
