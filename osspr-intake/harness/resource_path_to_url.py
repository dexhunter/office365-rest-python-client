"""Baseline for office365.runtime.paths.resource_path.ResourcePath.to_url.

A paged export of a large list asks for each item's resource URL. Every item
shares the collection path (sites / site-id / lists / list-id / items) and adds
its own id. Setup is excluded from the timed section.
Prints milliseconds for one pass.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.runtime.paths.resource_path import ResourcePath

# A list past the 5,000-item view threshold, read page by page.
ITEM_COUNT = 80_000


def _collection_path() -> ResourcePath:
    site = "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222"
    return ResourcePath(
        "items",
        ResourcePath("b9c1e4a0-7d2f-4c18-9a55-0e6f3d8b1c42", ResourcePath("lists", ResourcePath(site, ResourcePath("sites")))),
    )


def main() -> None:
    parent = _collection_path()
    items = [ResourcePath(str(i), parent) for i in range(ITEM_COUNT)]
    # One known address, so a broken builder fails before the measurement.
    expected = (
        "/sites/contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,"
        "22222222-2222-2222-2222-222222222222/lists/b9c1e4a0-7d2f-4c18-9a55-0e6f3d8b1c42/items/0"
    )
    if items[0].to_url() != expected:
        raise SystemExit(f"unexpected url: {items[0].to_url()}")
    items[0].to_url()  # warmup
    started = time.perf_counter()
    for item in items:
        item.to_url()
    elapsed_ms = (time.perf_counter() - started) * 1000
    print(f"{elapsed_ms:.4f}")


if __name__ == "__main__":
    main()
