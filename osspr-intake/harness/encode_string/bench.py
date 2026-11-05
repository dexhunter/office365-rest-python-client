"""Median time to OData-encode a block of server-relative file paths.

6,000 paths, the strings ``GetFileByServerRelativeUrl`` puts inside a path
literal: spaces and slashes, plus the characters the encoder rewrites
(quote, hash, question mark, ampersand, percent, plus, star). Warmup is not
included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.runtime.paths.builder import ODataPathBuilder

REPEATS = 21
WARMUP = 8

_PATHS = []
for _i in range(6000):
    if _i % 5 == 0:
        _PATHS.append(
            f"/sites/hr/Shared Documents/O'Brien & Co #{_i}/Q3 report {_i}?.docx"
        )
    elif _i % 5 == 1:
        _PATHS.append(f"/sites/finance/Reports/FY24/100%+complete/board pack {_i}*.pdf")
    else:
        _PATHS.append(
            f"/sites/hr/Shared Documents/2024/Q{_i % 4 + 1}/Clients/Acme Corp {_i}/Contracts/final.docx"
        )


def _block() -> None:
    for path in _PATHS:
        ODataPathBuilder._encode_string(path)


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
