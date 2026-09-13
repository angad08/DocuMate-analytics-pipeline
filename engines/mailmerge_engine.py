"""
The Word Mail Merge engine. Shared by X, Y and O.

X, Y and O each carried their own copy of this:

    Y   no batching,     reads Excel      about 175 lines
    X   batches of 250,  reads Excel      about 215 lines
    O   batches of 100,  reads database   about 215 lines

None of that is really an engine difference. Where the records came from
stopped mattering once both sources hand over the same list. And batching
is just a number. So: one engine, and batch_size says how many records go
into each run of Word. batch_size = None does what Y did.


WHY WE HAND WORD A .CSV AND NOT THE WORKBOOK
--------------------------------------------
Merging straight from the .xlsx made Word pop up its "Select Table" box
every single run.

The cause is the Excel driver (ACE OLEDB). Give Word a Connection string
and it treats the workbook as a database, and in a database every sheet is
a table - so Word asks which table before it will run your SQLStatement.
The SQLStatement was fine, it just arrived too late to matter.

Things tried, and ruled out:
  - the data source saved inside the template
  - setting MainDocumentType before OpenDataSource
  - the idea that it was confused by having several sheets: a temp
    workbook with exactly ONE sheet still opened the picker, and listed
    that one sheet
  - SubType 0, 1 and 8, and leaving SQLStatement out completely

Every Excel-based version got stuck. A .csv has only one table in it, so
there is nothing to pick and nothing to ask - it was the only thing that
connected on its own, in under five seconds.

So: no Connection, no SQLStatement, just the .csv filename. Putting a
Connection string back would bring the dialog back with it.

Writing only the pending rows into that .csv also killed the old "records
must be in one unbroken run" rule. Gaps only mattered because the merge ran
against the whole sheet with PRINTED rows scattered through it. This file
holds only the rows we want, so the range is always 1 to N.
"""

import os
import shutil
import time

import pandas as pd

from setup import messages
from engines import word_com
from engines.output_paths import delete_temp_file
from engines.output_paths import make_output_folder


# How many records go into each run of Word. Word stops responding if you
# ask it to build one enormous document in a single go - at 1,440 records
# it hung completely.
DEFAULT_BATCH_SIZE = 250

# Word needs a moment between COM calls. Take these out and the merge
# starts failing now and then on slower machines.
SETTLE = 0.5
STARTUP = 1.0

TEMP_CSV_NAME = "__temp_merge_source.csv"


def plan_batches(total, batch_size):
    """
    Work out the record ranges to give Word, e.g. (1, 250), (251, 500)...

    Both ends are included. If batch_size is None, or bigger than the
    number of records, you get one batch covering everything - which is
    what Y did.

    This is plain arithmetic, no Word involved, so it is the part that can
    be tested properly. It is also the part that would silently skip or
    duplicate a certificate if it were wrong.
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
    Write the pending records to the temp .csv that Word will merge from.

    utf-8-sig matters: Word needs that marker at the start of the file to
    read the header row properly.
    """
    path = os.path.join(os.path.abspath(folder), TEMP_CSV_NAME)
    pd.DataFrame(list(records)).to_csv(path, index=False, encoding="utf-8-sig")
    return path


class MailMergeEngine:
    """Drives Word's own Mail Merge through COM."""

    # Shown in messages.
    label = "Word Mail Merge"

    def __init__(self, template_path, batch_size=DEFAULT_BATCH_SIZE, visible=True):
        self.template_path = os.path.abspath(template_path)
        self.batch_size = batch_size
        self.visible = visible

    def generate(self, records, output_path):
        """Merge every record and save one document."""

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

            # DispatchEx, not Dispatch. Dispatch would attach to a copy of
            # Word the user already has open, and whatever state that copy
            # is in would leak into our merge. DispatchEx starts a fresh one.
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

            # No Connection and no SQLStatement on purpose. See the notes at
            # the top of this file - either one brings back the dialog.
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

            # --- run each batch ---
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

            # Close the template before joining, so Word isn't holding it.
            self.close_document(template)
            template = None

            combined = self.join_batches(word, batch_files, output_path)

            print(messages.SAVED.format(count=total, path=output_path))
            return output_path

        except Exception as error:
            print(messages.MM_ERROR.format(error=error))
            raise

        finally:
            self.close_document(combined)
            self.close_document(template)
            self.quit_word(word)

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            # The .csv holds real personal data, so it must not survive the
            # run - whether the merge worked or not.
            for path in batch_files + [csv_path]:
                delete_temp_file(path)

    def run_one_batch(self, word, template, first, last, folder, number):
        """Run Word once for one range of records, save the result as a temp file."""

        template.MailMerge.DataSource.FirstRecord = int(first)
        template.MailMerge.DataSource.LastRecord = int(last)
        template.MailMerge.Execute(Pause=False)
        time.sleep(SETTLE)

        result = word.ActiveDocument

        path = os.path.join(os.path.abspath(folder), "__temp_batch_" + str(number) + ".docx")
        result.SaveAs2(path, FileFormat=word_com.FORMAT_DOCX)

        self.close_document(result)
        return path

    def join_batches(self, word, batch_files, output_path):
        """
        Join the batch files into one document.

        If there is only one batch we just move it. Opening and re-saving it
        would only risk Word reflowing the layout for no reason.
        """
        if len(batch_files) == 1:
            shutil.move(batch_files[0], output_path)
            del batch_files[:]
            return None

        combined = word.Documents.Open(batch_files[0])

        for path in batch_files[1:]:
            end = combined.Content.End - 1
            combined.Range(end, end).InsertBreak(word_com.SECTION_BREAK_NEXT_PAGE)

            end = combined.Content.End - 1
            combined.Range(end, end).InsertFile(path)

        combined.SaveAs2(output_path, FileFormat=word_com.FORMAT_DOCX)
        return combined

    def close_document(self, document):
        """Close a document, ignoring the case where it's already gone."""
        if document is None:
            return
        try:
            document.Close(word_com.DO_NOT_SAVE_CHANGES)
        except Exception:
            pass

    def quit_word(self, word):
        """Shut Word down, ignoring the case where it's already gone."""
        if word is None:
            return
        try:
            word.Quit()
        except Exception:
            pass
