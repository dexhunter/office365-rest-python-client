"""Byte size of record_to_json over a migration batch of list rows.

Each row is one persisted record (the bundle writer calls record_to_json once
per item). The reported value is the sum of those UTF-8 payloads.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.runtime.converters.json_file import record_to_json

ROWS = 4000


def _rows(n: int) -> list[dict]:
    rows = []
    for i in range(n):
        rows.append(
            {
                "Title": f"Invoice {i}",
                "Amount": round(i * 1.25, 2),
                "Owner": f"user{i}@contoso.com",
                "Status": "Active" if i % 2 == 0 else "Closed",
                "Created": "2024-06-01T00:00:00Z",
                "Department": "Finance",
                "CostCenter": f"CC-{i % 40:03d}",
                "Notes": "Imported from the month-end workbook.",
                "Approved": i % 3 == 0,
                "LineItems": [{"sku": f"SKU-{j}", "qty": j + 1} for j in range(3)],
            }
        )
    return rows


def main() -> None:
    rows = _rows(ROWS)
    total = 0
    for row in rows:
        text = record_to_json(row)
        total += len(text.encode("utf-8"))
    parsed = json.loads(record_to_json(rows[0]))
    if parsed["Title"] != rows[0]["Title"] or parsed["LineItems"] != rows[0]["LineItems"]:
        raise SystemExit("record_to_json did not round-trip")
    print(total)


if __name__ == "__main__":
    main()
