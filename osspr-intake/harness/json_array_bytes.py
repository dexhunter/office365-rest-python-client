"""Baseline output size for JsonArrayStreamWriter.write.

One streamed JSON array page: 2500 records, 22 short columns, the shape of a
wide list or user export. Prints the UTF-8 byte length of the file the writer
closes (opening '[', records, closing ']').
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.runtime.converters.streamers import JsonArrayStreamWriter

ROWS = 2500
FIELDS = 20


def main() -> None:
    records = []
    for index in range(ROWS):
        row = {f"F{field}": "x" for field in range(FIELDS)}
        row["Id"] = index
        row["Title"] = f"T{index}"
        records.append(row)
    buffer = io.StringIO()
    writer = JsonArrayStreamWriter(buffer)
    writer.write(records)
    writer.close()
    print(len(buffer.getvalue().encode("utf-8")))


if __name__ == "__main__":
    main()
