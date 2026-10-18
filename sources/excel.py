"""
The Excel data source. Shared by v3, X and Y.

Reading uses pandas. Writing back uses openpyxl on purpose - pandas would
rewrite the whole sheet and wipe out your formulas, colours and column
widths. openpyxl edits the cells and leaves the rest alone.
"""

import os
import time
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

from setup import config
from setup import messages
from sources import excel_format


class ExcelSource:
    """Reads records from a worksheet, and writes statuses back to it."""

    # Shown in messages, e.g. "Reading Excel..."
    label = "Excel"

    def __init__(self, path, sheet):
        self.path = path
        self.sheet = sheet

    # -- reading ------------------------------------------------------------

    def validate_source(self):
        """
        Stop now if the file or the sheet isn't there.

        The sheet used to go unchecked, so a renamed or deleted one only
        showed up as a pandas error once the run was already under way.
        """
        if not os.path.exists(self.path):
            raise FileNotFoundError(messages.SOURCE_MISSING.format(path=self.path))

        # The sheet itself is checked in load_records(), where it is read.
        # Opening the workbook twice would double the cost of a large file
        # and report a locked file from the wrong place.

    def load_records(self):
        """
        Read the sheet, and put it in certificate shape if it isn't already.

        A sheet that was built with the old Excel formulas is used exactly
        as it is. A raw entry sheet is converted here instead - see
        sources/excel_format.py, which holds the rules those formulas used
        to hold.
        """
        try:
            frame = pd.read_excel(self.path, sheet_name=self.sheet)
        except ValueError as error:
            if "not found" not in str(error).lower():
                raise
            sheets = ", ".join(pd.ExcelFile(self.path).sheet_names)
            raise ValueError(messages.EXCEL_NO_SHEET.format(
                sheet=self.sheet, path=self.path, sheets=sheets))

        return excel_format.prepare(frame)

    def close(self):
        """Nothing to close for a file. Here so both sources match."""
        pass

    # -- writing back -------------------------------------------------------

    def mark_printed(self, data):
        """Mark these records as PRINTED. Kept for callers that want it."""
        return self.mark_status(data, config.DONE_STATUS)

    def mark_status(self, data, status_value):
        """
        Mark only rows whose Serial is in data as PRINTED, and stamp today's date.
        """
        try:
            workbook = load_workbook(self.path)
            sheet = workbook[self.sheet]

            # The same sheet again, showing what each formula displays instead
            # of the formula. Serial is often a formula (="...A2 & "/" & ..."),
            # and pandas read the displayed value, so that is what has to be
            # matched. Only ever read from this copy - saving it would replace
            # every formula with its last value.
            shown = load_workbook(self.path, data_only=True)[self.sheet]

            # Map column name -> column number, from the header row.
            columns = {}
            for number, cell in enumerate(sheet[1], start=1):
                columns[cell.value] = number

            if "STATUS" not in columns:
                raise ValueError(messages.EXCEL_NO_STATUS_COLUMN)

            # A raw entry sheet may call the column "S No.", and holds the
            # bare number: the "/year" was added on read, from the
            # registration date. It is rebuilt the same way here so the two
            # can be matched.
            serial_name = None
            for name in excel_format.RAW_ALIASES["S No."]:
                if name in columns:
                    serial_name = name
            if serial_name is None:
                raise ValueError(messages.EXCEL_NO_SERIAL_COLUMN)
            registration_column = columns.get(
                excel_format.RAW_COLUMNS["registration_date"])

            # Older sheets don't have Date_Printed, so add it.
            if "Date_Printed" not in columns:
                new_column = sheet.max_column + 1
                sheet.cell(row=1, column=new_column, value="Date_Printed")
                columns["Date_Printed"] = new_column

            status_column = columns["STATUS"]
            date_column = columns["Date_Printed"]
            serial_column = columns[serial_name]
            serials = set(data["Serial"])
            today = datetime.now().strftime("%d/%m/%Y")
            marked = 0

            for row in range(2, sheet.max_row + 1):
                serial = shown.cell(row=row, column=serial_column).value
                if serial not in serials and registration_column:
                    serial = excel_format.serial_for(
                        serial,
                        shown.cell(row=row, column=registration_column).value,
                    )
                if serial not in serials:
                    continue
                status = sheet.cell(row=row, column=status_column).value
                if status is None:
                    status = ""

                # Never move a record backwards. A row already PRINTED is
                # finished, and a row already at the status being written
                # has nothing to change.
                current = str(status).strip().upper()
                if current == config.DONE_STATUS or current == status_value:
                    continue

                sheet.cell(row=row, column=status_column, value=status_value)
                sheet.cell(row=row, column=date_column, value=today)
                marked = marked + 1

            workbook.save(self.path)

            if marked != len(data):
                print(messages.EXCEL_COUNT_WARNING.format(expected=len(data), marked=marked))

            print(messages.EXCEL_UPDATED.format(
                sheet=self.sheet,
                date=today,
                clock=time.strftime("%H:%M:%S"),
            ))

        except PermissionError:
            print(messages.EXCEL_LOCKED)
            raise

        except Exception as error:
            print(messages.EXCEL_UPDATE_ERROR.format(error=error))
            raise
