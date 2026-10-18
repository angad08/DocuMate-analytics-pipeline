"""
Asking the user, and marking the records as completed.

Both halves of that live here together, because they are one decision:
nothing is ever marked completed unless the user says yes, so the question
and the update belong in the same place.

Where the update actually goes depends on the version - the Excel versions
write back to the sheet, the database versions run an UPDATE. That part is
in sources/excel.py and sources/postgres.py. This file just asks, then
hands the records over.
"""

import time

from setup import config
from setup import messages
from setup import ui


def update_status(sink, data, count, failed=0):
    """
    Ask whether to mark these records as COMPLETED, and do it if yes.

    sink   - where the statuses get written (the Excel sheet, or the database)
    data   - the records that were just turned into documents
    count  - how many succeeded, for the question
    failed - how many failed to render

    Returns (text, seconds): a bit of text for the summary popup and how
    long the write itself took. The seconds are the write only, not the
    wait for an answer - a run should not look slower because somebody
    took a moment to read the question.

    The question is always asked, and the answer always wins over whatever
    the version was set up with. That was true in the original scripts and
    it stays true here.
    """
    if not ui.confirm(messages.ASK_UPDATE.format(
        count=count,
        failed=" ({} failed)".format(failed) if failed else "",
    )):
        return "", 0.0

    print(messages.UPDATING.format(target=sink.label))
    started = time.time()

    # COMPLETED, not PRINTED: the documents exist, but nothing has been
    # sent to a printer yet. Printing is offered next and promotes these to
    # PRINTED if it happens.
    sink.mark_status(data, config.COMPLETED_STATUS)

    seconds = time.time() - started

    return messages.UPDATE_SUFFIX.format(
        target=sink.label,
        seconds=seconds,
    ), seconds
