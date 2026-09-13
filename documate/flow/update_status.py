"""
Asking the user, and marking the records as printed.

Both halves of that live here together, because they are one decision:
nothing ever gets marked printed unless the user says yes, so the question
and the update belong in the same place.

Where the update actually goes depends on the version - the Excel versions
write back to the sheet, the database versions run an UPDATE. That part is
in sources/excel.py and sources/postgres.py. This file just asks, then
hands the records over.
"""

import time

from documate.setup import messages
from documate.setup import ui


def update_status(sink, data, count):
    """
    Ask whether to mark these records as PRINTED, and do it if yes.

    sink   - where the statuses get written (the Excel sheet, or the database)
    data   - the records that were just turned into documents
    count  - how many, for the question

    Returns a bit of text for the summary popup, or an empty string if the
    user said no.

    The question is always asked, and the answer always wins over whatever
    the version was set up with. That was true in the original scripts and
    it stays true here.
    """
    if not ui.confirm(messages.ASK_UPDATE.format(count=count)):
        return ""

    print(messages.UPDATING.format(target=sink.label))
    started = time.time()

    sink.mark_printed(data)

    return messages.UPDATE_SUFFIX.format(
        target=sink.label,
        seconds=time.time() - started,
    )
