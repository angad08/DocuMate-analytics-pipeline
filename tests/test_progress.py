"""
The progress bar, as a terminal would show it.

Only a terminal redraws the bar in place, so these tests write to a fake
one and replay the carriage returns to see what would be left on screen.
"""

import io

from setup import messages
from setup import ui


class Terminal(io.StringIO):
    def isatty(self):
        return True

    def screen(self):
        """The lines on screen: each \\r goes back and types over the line."""
        lines = []
        for raw in self.getvalue().split("\n"):
            line = ""
            for part in raw.split("\r"):
                line = part + line[len(part):]
            lines.append(line.rstrip())
        return lines


def use_terminal(monkeypatch):
    """
    Send output to a fake terminal. Called in the test itself: pytest puts
    its own capture back on sys.stdout after fixtures have run.
    """
    screen = Terminal()
    monkeypatch.setattr("sys.stdout", screen)
    monkeypatch.setattr(ui, "_printed_width", 0)
    monkeypatch.setattr(ui, "_open_bar", None)
    monkeypatch.setattr(ui, "_characters", ("#", "."))
    return screen


def test_batch_lines_do_not_pick_up_the_end_of_the_bar(monkeypatch):
    terminal = use_terminal(monkeypatch)
    total = 1556
    batches = [(1, 250), (251, 500), (501, 750), (751, 1000),
               (1001, 1250), (1251, 1500), (1501, 1556)]

    for number, (first, last) in enumerate(batches, start=1):
        ui.say(messages.MM_BATCH.format(number=number, total=len(batches),
                                        first=first, last=last))
        ui.progress_bar(last, total, messages.RECORD_PROGRESS)

    screen = [line for line in terminal.screen() if line]
    assert screen[:7] == [
        messages.MM_BATCH.format(number=number, total=7, first=first, last=last)
        for number, (first, last) in enumerate(batches, start=1)
    ]
    assert len(screen) == 8
    assert screen[7].endswith("1556/1556 records")


def test_say_with_no_bar_is_an_ordinary_line(monkeypatch):
    terminal = use_terminal(monkeypatch)
    ui.say("DocuMate : hello")
    assert terminal.getvalue() == "DocuMate : hello\n"
