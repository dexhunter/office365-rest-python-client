"""Median time to serialize a block of scalar cells through ``serialize_value``.

32,000 values, the mix of a ``$select`` page of scalar columns: strings, ints,
bools, and nulls. The profile object is built once, outside the timer.
Warmup is not included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.runtime.converters.value import serialize_value
from office365.runtime.odata.v4.json_format import V4JsonFormat

REPEATS = 21
WARMUP = 8

_PROFILE = V4JsonFormat()
_CELLS: list[object] = []
for _i in range(4000):
    _CELLS.extend((f"Task {_i}", _i, True, None, f"user{_i}@contoso.com", "Retail", 12, False))


def _block() -> None:
    for value in _CELLS:
        serialize_value(value, _PROFILE)


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
