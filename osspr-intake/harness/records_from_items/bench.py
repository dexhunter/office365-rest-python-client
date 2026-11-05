"""Median time to project one loaded SharePoint list page into records.

120 items, each with the built-in columns plus 40 custom site columns, the
shape of a paged ``RenderListDataAsStream`` / items read. Hydration happens
once, before the timer. An empty ``$select`` makes the projection ask the
type for its columns, which is what an export does when the caller did not
narrow the query. Warmup is not included. ``gc.collect`` runs between repeats,
outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.runtime.converters.records import records_from_items
from office365.runtime.odata.v3.json_light_format import JsonLightFormat
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.collection import ListItemCollection

REPEATS = 21
WARMUP = 8

_CUSTOM = [f"CustomCol{k}" for k in range(40)]
_PAGE = {
    "d": {
        "results": [
            {
                "__metadata": {"type": "SP.Data.TasksListItem", "etag": '"8"'},
                "Id": i,
                "Title": f"Task {i}",
                "Created": "2023-04-11T09:22:18.1234567Z",
                "Modified": "2025-10-02T18:03:55.9876543Z",
                "AuthorId": 7,
                "EditorId": 12,
                "ContentTypeId": "0x010100C5",
                **{name: f"value {i} {name}" for name in _CUSTOM},
            }
            for i in range(1, 121)
        ]
    }
}


def _load():
    context = ClientContext("https://contoso.sharepoint.com")
    collection = ListItemCollection(context)
    JsonLightFormat().hydrate(collection, _PAGE)
    return list(collection)


_ITEMS = _load()


def _block() -> None:
    records_from_items(_ITEMS, [], [])


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
