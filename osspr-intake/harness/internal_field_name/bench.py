"""Median time to sanitize column names for one dataframe import chunk.

``records_from_dataframe`` calls ``internal_field_name`` once per cell, and a
chunk is ``SAFE_PAGE_SIZE`` (2000) rows. These 25 headers are the display names
a task-list CSV actually has: spaces, punctuation, and a few titles that
collide with built-in fields. The header list is built once, before the timer.
Every cell of a column repeats the same name, and each call still runs a
regular expression over the non-word characters.
Warmup is not included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.sharepoint.fields.name import internal_field_name

REPEATS = 21
WARMUP = 8

_ROWS = 2000
_COLUMNS = (
    "Title",
    "Name",
    "Q3 Revenue (USD)",
    "Owner / Name",
    "Status",
    "Due Date",
    "Customer E-mail",
    "Region-East",
    "Cost Center #",
    "Project Code",
    "Priority",
    "Percent Complete",
    "Assigned To",
    "Start Date",
    "Category",
    "Department",
    "External Id",
    "Notes",
    "Approval Status",
    "Fiscal Year",
    "Vendor Name",
    "PO Number",
    "Site URL",
    "Created By",
    "Modified",
)


def _block() -> None:
    for _row in range(_ROWS):
        for name in _COLUMNS:
            internal_field_name(name)


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
