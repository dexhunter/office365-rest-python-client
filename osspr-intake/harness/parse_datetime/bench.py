"""Median time to parse a block of Graph and SharePoint datetime literals.

The block is 4,900 stamps, the mix a list sync actually receives: UTC ``Z``,
seven-digit .NET fractions, numeric offsets, and date-times with no zone.
Warmup is not included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.runtime.converters.scalars import parse_datetime

REPEATS = 21
WARMUP = 8

# One stamp family per line, repeated. 7 * 700 = 4,900 parses.
_STAMPS = (
    "2023-04-11T09:22:18Z",
    "2025-10-02T18:03:55.9876543Z",
    "2024-08-19T16:42:01.5000000Z",
    "2025-01-06T04:12:44.2500000+00:00",
    "2022-12-31T23:59:59.1234567-08:00",
    "2019-11-02T00:00:00",
    "2020-03-14T08:15:30.1234567Z",
) * 700


def _block() -> None:
    for stamp in _STAMPS:
        parse_datetime(stamp)


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
