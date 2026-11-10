"""Baseline for office365.runtime.converters.value.to_snake_case.

A page of list items repeats the same undeclared camelCase column names.
ClientObject.get_property resolves each of those with to_snake_case.
Prints milliseconds for ROWS passes over NAMES, setup excluded.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.runtime.converters.value import to_snake_case

ROWS = 5000
NAMES = (
    "Title",
    "Status",
    "Priority",
    "CustomColumn",
    "OwnerEmail",
    "DueDate",
    "Amount",
    "IsArchived",
    "Department",
    "CostCenter",
    "ProjectCode",
    "Notes",
    "Category",
    "ModifiedBy",
    "CreatedBy",
)


def main() -> None:
    names = NAMES
    started = time.perf_counter()
    for _ in range(ROWS):
        for name in names:
            to_snake_case(name)
    elapsed_ms = (time.perf_counter() - started) * 1000
    print(f"{elapsed_ms:.4f}")


if __name__ == "__main__":
    main()
