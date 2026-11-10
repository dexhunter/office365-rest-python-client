"""Baseline for office365.formats.text.fold_line.

One long content line, the shape of an iCalendar DESCRIPTION or a vCard NOTE
(meeting minutes). encode() and ContactRecord.vcard() fold every line this way.
Prints milliseconds for REPEATS calls, setup excluded.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.formats.text import fold_line

REPEATS = 25
_SENTENCE = "Quarterly review notes for the Contoso migration, including cafe owners and dates. "
_LINE = "DESCRIPTION:" + (_SENTENCE * 700)


def main() -> None:
    line = _LINE
    started = time.perf_counter()
    for _ in range(REPEATS):
        fold_line(line)
    elapsed_ms = (time.perf_counter() - started) * 1000
    print(f"{elapsed_ms:.4f}")


if __name__ == "__main__":
    main()
