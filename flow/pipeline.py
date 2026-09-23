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

from setup import messages
from setup import ui
from flow import checks
from engines.output_paths import build_output_name
from flow.update_status import update_status
from flow.checks import CheckFailed


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
                update_text = update_status(self.sink, printed_data, printed_count,
                                            failed=failed_count)
            except Exception as error:
                ui.notify(messages.STATUS_UPDATE_FAILED.format(
                    path=output_path,
                    source=self.sink.label,
                    reason=error,
                ))
                return

            ui.notify(messages.SUCCESS.format(
                count=printed_count,
                failed_text=messages.RENDER_FAILURES.format(failed=failed_count)
                            if failed_count else "",
                update_text=update_text,
                seconds=time.time() - started,
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

    def check_records(self, data):
        """
        Run the v3 checks, if this version uses them.

        Z and O filter in their SQL and have checks turned off, which
        matches how they always worked. Flip check_records to True in the
        version list if you ever want them on.
        """
        try:
            if self.version.check_records:
                return checks.check_records(data)

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

        return checks.sort_by_serial(data)

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

        print(messages.MERGE_DONE.format(count=len(successful)))
        print(messages.MERGE_TIME.format(
            engine=self.engine.label,
            seconds=time.time() - started,
        ))
        return successful, output_path

    # -- running more than once ---------------------------------------------

    def start(self, poll=False, interval=300):
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
