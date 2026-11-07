"""Median time to reduce one page of message bodies to comparable text.

30 HTML bodies, each about 6 KB: a head with a style and a script, comments,
and 40 paragraphs of the markup Graph stores on an event or a message. The
bodies are built once, before the timer. A migration digest compares
``ItemBody.content`` with ``Compare(format="text")``, which runs this helper.
It strips head, comments, and tags in separate regular-expression passes, then
unescapes and collapses whitespace.
Warmup is not included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.runtime.odata.comparable import _text

REPEATS = 21
WARMUP = 8

_HTML = "<html><head><style>.x{color:red}</style><script>var a=1;</script></head><body>" + "".join(
    "<p>Paragraph "
    f"{n} of the quarterly review with <b>budget</b> and "
    f"<a href='https://contoso.sharepoint.com/sites/finance/q{n}'>the working paper</a>.</p>"
    f"<!-- note {n} -->"
    for n in range(40)
) + "</body></html>"
_BODIES = [_HTML.replace("quarterly", f"quarterly {i}") for i in range(30)]


def _block() -> None:
    for body in _BODIES:
        _text(body)


def main() -> None:
    for _ in range(WARMUP):
        _block()
    samples: list[float] = []
    for _ in range(REPEATS):
        gc.collect()
        started = time.perf_counter()
        _block()
        samples.append((time.perf_counter() - started) * 1000.0)
    print(f"median_ms={statistics.median(samples):.4f}")


if __name__ == "__main__":
    main()
