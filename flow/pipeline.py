"""
The steps every version runs, in order.

All five versions did the same things:

    validate the template file -> load the records -> check them
    -> transform them -> make the documents -> update the status

The old run() methods differed only in which steps they skipped and what
they printed. Both of those are settings now, read from the version, so
there is one copy of the sequence.

The trick that makes it work: this file never mentions Excel, the database,
docxtpl or Word. It only ever says self.source and self.engine, and those
get plugged in when the version starts.
"""

import os
import time

import pandas as pd

from setup import messages
from setup import ui
from flow import checks
from engines.output_paths import build_output_name
from flow.update_status import update_status
from flow.checks import CheckFailed
from setup import config


class Pipeline:
    """Runs one version from start to finish."""

    def __init__(self, version, source, engine, sink=None):
        self.version = version
        self.source = source
        self.engine = engine

        # Most versions write statuses back to wherever they read from.
        if sink is None:
            sink = source
        self.sink = sink

    # -- one run ------------------------------------------------------------

    def run(self):
        """
        Do the whole thing once.

        This returns normally whether it worked, was stopped by a failed
        check, or hit an error. Errors get shown to the user rather than
        thrown, because this is the top level of a desktop tool - there is
        nobody above us to catch them.
        """
        print(messages.BANNER.format(version=self.version.label))
        print(messages.READY)

        try:
            self.validate_template_file()
            self.source.validate_source()

            print(messages.READING.format(source=self.source.label))
            data = self.source.load_records()

            print(messages.SEARCHING)
            data = self.check_records(data)
            if data is None:
                return          # a check failed, message already shown

            data = self.transform_records(data)
            count = len(data)
            print(messages.FOUND.format(count=count))

            started = time.time()

            successful, output_path = self.make_documents(data)
            if not successful:
                raise RuntimeError(messages.NOTHING_TO_MERGE)

            printed_data = data.iloc[[number - 1 for number in successful]]
            printed_count = len(printed_data)
            failed_count = count - printed_count
            try:
                update_text, write_seconds = update_status(
                    self.sink, printed_data, printed_count,
                    failed=failed_count)
            except Exception as error:
                ui.notify(messages.STATUS_UPDATE_FAILED.format(
                    path=output_path,
                    source=self.sink.label,
                    reason=error,
                ))
                return

            # Added up from the parts rather than read off a clock. A wall
            # clock counts the PDF, the printing, and the seconds somebody
            # spends reading a question before answering it - none of which
            # is work DocuMate did.
            elapsed = self.generate_seconds() + write_seconds

            self.offer_printing(printed_data, printed_count, output_path)

            ui.notify(messages.SUCCESS.format(
                count=printed_count,
                failed_text=messages.RENDER_FAILURES.format(failed=failed_count)
                            if failed_count else "",
                update_text=update_text,
                seconds=elapsed,
                engine=self.engine.label,
                source=self.source.label,
            ))

        except PermissionError:
            # The spreadsheet or the output file is open in Excel or Word.
            ui.notify(messages.FILE_LOCKED)

        except Exception as error:
            ui.notify(messages.UNEXPECTED.format(error=error))

    # -- the steps ----------------------------------------------------------

    def validate_template_file(self):
        """
        Stop straight away if the Word template isn't there.

        Only Z and O used to check this. Doing it for every version means a
        typo in a path costs you a second, instead of showing up after the
        sheet has been read and Word is already open on screen.
        """
        template = self.version.template_path()

        if not os.path.exists(template):
            raise FileNotFoundError(messages.TEMPLATE_MISSING.format(path=template))

    def fix_serials(self, data, missing):
        """
        Decide what to do about blank S No. cells.

        Asks when someone is at the keyboard, and numbers every row by its
        position for that run only if they say yes - the workbook is never
        written to, so the column still has to be filled in properly later.
        When polling there is nobody to ask, so it stops and says why.
        """
        if getattr(self, "polling", False):
            raise CheckFailed(messages.SERIAL_UNATTENDED.format(count=len(missing)))

        if not ui.confirm(messages.SERIAL_ASK.format(
                count=len(missing), total=len(data))):
            raise CheckFailed(messages.SERIAL_REFUSED.format(count=len(missing)))

        data = checks.renumber_serials(data)
        print(messages.SERIAL_RENUMBERED.format(total=len(data)))
        return data

    def check_template_fields(self, data):
        """
        Stop if the template asks for a field the rows cannot fill, and
        return the columns it prints - the ones that must be filled in.

        Returns None for the birth register, whose own rules decide which
        columns must be filled. Also None when there is no template to read
        (a test's stand-in version), so nothing is checked that cannot be.
        """
        if data is None or not hasattr(self.version, "template_path"):
            return None

        from setup import picker

        fields = picker.template_fields(self.version.template_path(),
                                        self.version.engine)
        missing = picker.missing_fields(fields, data.columns)
        if missing:
            raise CheckFailed(messages.TEMPLATE_FIELDS_MISSING.format(
                count=len(missing),
                fields=", ".join(missing),
                columns=", ".join(str(c) for c in data.columns
                                  if c not in config.STATE_COLUMNS_OPTIONAL),
            ))

        if config.is_register():
            return None
        return picker.columns_for_fields(fields, data.columns)

    def allow_blanks(self, rows):
        """Ask whether blank printed fields may stay blank. Never when polling."""
        if getattr(self, "polling", False):
            return False
        return ui.confirm(messages.BLANKS_ASK.format(count=rows))

    def check_records(self, data):
        """
        Run the v3 checks, if this version uses them.

        Z and O filter in their SQL and have checks turned off, which
        matches how they always worked. Flip check_records to True in the
        version list if you ever want them on.
        """
        try:
            required = self.check_template_fields(data)

            if self.version.check_records:
                # The date columns go in so an unreadable date is reported as a
                # format problem rather than a missing value: 8.11.26 is plainly
                # not an empty cell, and saying "missing" sends the officer to
                # look at one.
                return checks.check_records(data, self.fix_serials,
                                            self.version.date_columns or (),
                                            required=required,
                                            blanks_ok=self.allow_blanks)

            # Even with checks off, an empty result still stops the run.
            if data is None or data.empty:
                raise CheckFailed(messages.NO_DATA)

            return data

        except CheckFailed as failure:
            for line in failure.extra_lines:
                print(line)
            ui.notify(failure.message)
            return None

    def transform_records(self, data):
        """Format the dates and put the records in printing order."""
        if self.version.date_columns:
            data = checks.format_dates(data, self.version.date_columns)

        # Any other column Excel holds as dates would print with a time on
        # it - "2026-01-05 00:00:00" - so it gets the same format. This is
        # what lets a case study's dates come out right without listing
        # anyone having to say which columns are dates.
        stamped = [name for name in data.columns
                   if pd.api.types.is_datetime64_any_dtype(data[name])]
        if stamped:
            data = checks.format_dates(data, stamped)

        return checks.sort_by_serial(data)

    def offer_printing(self, data, count, output_path):
        """
        Offer to print the merged document, and mark what was printed.

        Everything here is optional and everything fails softly. The
        documents are already saved by the time this runs, so no answer and
        no failure can lose work - the worst case is that the records stay
        as they are and the file is printed by hand later.
        """
        if not config.offer_printing():
            return

        from engines import printing

        try:
            printers, default = printing.available_printers()
        except RuntimeError as error:
            print(messages.PRINTING_UNAVAILABLE.format(error=error))
            return

        if not printers:
            print(messages.NO_PRINTERS)
            return

        if not ui.confirm(messages.ASK_PRINT.format(count=count)):
            print(messages.PRINT_SKIPPED.format(path=output_path))
            return

        chosen = self.choose_printer(printers, default)
        if chosen is None:
            print(messages.PRINT_SKIPPED.format(path=output_path))
            return

        # Asked before the job is sent, while there is still something
        # useful to say about it.
        problem = printing.printer_problem(chosen["name"])
        if problem:
            ui.notify(messages.PRINT_PROBLEM.format(problem=problem))
            return

        print(messages.PRINT_SENDING.format(count=count, printer=chosen["name"]))

        try:
            printing.print_document(output_path, chosen["name"])
        except RuntimeError as error:
            ui.notify(messages.PRINT_PROBLEM.format(problem=error))
            return

        print(messages.PRINT_SENT)

        # Only now do the records become PRINTED. A record that was never
        # marked COMPLETED moves straight here from IN PROCESS, which is
        # what happens when the status question was answered no and the
        # print question yes.
        try:
            self.sink.mark_status(data, config.DONE_STATUS)
        except Exception as error:
            ui.notify(messages.STATUS_UPDATE_FAILED.format(
                path=output_path,
                source=self.sink.label,
                reason=error,
            ))
            return

        print(messages.PRINT_NOT_CONFIRMED)

    def choose_printer(self, printers, default):
        """
        Show the printers and return the one chosen, or None to skip.

        The default is offered first so the usual answer is Enter. Printers
        that write a file rather than paper are marked, because on a lot of
        machines the default is Microsoft Print to PDF and sending a
        hundred certificates there looks like it worked.
        """
        lines = []
        default_index = None

        for number, printer in enumerate(printers):
            line = printer["name"]
            if printer["name"] == default:
                line = line + messages.PRINTER_DEFAULT_TAG
                default_index = number
            if printer["virtual"]:
                line = line + messages.PRINTER_VIRTUAL_TAG
            lines.append(line)

        print(messages.PRINTERS_HEADING)
        chosen = ui.choose(messages.CHOOSE_PRINTER, lines, default_index)

        if chosen is None:
            return None
        return printers[chosen]

    def make_documents(self, data):
        """Hand the records to the engine and report how long it took."""
        print(messages.ENGINE_START.format(engine=self.engine.label))
        started = time.time()

        records = checks.to_records(data)

        output_path = build_output_name(
            self.version.output_folder(),
            self.version.output_prefix,
            self.version.timestamp_format,
        )

        successful = self.engine.generate(records, output_path)

        # The engine prints what it made, how long it took and where it
        # went, in one line. Saying it again here only made the same
        # sentence appear twice.
        return successful, output_path

    def generate_seconds(self):
        """How long the engine spent making the documents."""
        return getattr(self.engine, "generate_seconds", 0.0)

    # -- running more than once ---------------------------------------------

    def start(self, poll=False, interval=300):
        # Remembered so the checks know whether anyone is watching. A
        # question asked inside a polling loop would wait for an answer
        # that is never coming.
        self.polling = poll

        """
        Run once, or keep checking for new records until you stop it.

        Only Z and O offered the checking loop, but there is nothing
        database-specific about it - it just calls run() again. So any
        version can use it now.
        """
        if not poll:
            try:
                self.run()
            finally:
                self.source.close()
            return

        print(messages.POLL_ENABLED)
        print(messages.POLL_INTERVAL.format(seconds=interval))

        try:
            while True:
                try:
                    self.run()
                    time.sleep(interval)

                except KeyboardInterrupt:
                    print(messages.POLL_STOPPED)
                    break

                except Exception as error:
                    print(messages.POLL_ERROR.format(error=error))
                    time.sleep(interval)
        finally:
            self.source.close()
