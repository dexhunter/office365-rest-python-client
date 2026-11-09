"""Time Value.to_xml by rendering a large CAML <In> of lookup ids.

SharePoint list queries filter a page of ids with <In><Values> under the list
view threshold. Every integer becomes its own <Value> element.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.sharepoint.listitems.caml.expressions import In
from office365.sharepoint.listitems.caml.fields import FieldRef
from office365.sharepoint.listitems.caml.values import Value

REPEATS = 250
IDS = 2500


def main() -> None:
    expr = In(FieldRef("Category", lookup_id=True), [Value(i) for i in range(1, IDS + 1)])
    xml = expr.to_xml()
    if xml.count("<Value ") != IDS:
        raise SystemExit(f"expected {IDS} Value elements")
    if not xml.startswith('<In><FieldRef Name="Category" LookupId="TRUE"/>'):
        raise SystemExit("In/FieldRef wrapper changed")
    if "<Value Type=\"Integer\">1</Value>" not in xml or f"<Value Type=\"Integer\">{IDS}</Value>" not in xml:
        raise SystemExit("integer values were not rendered")

    t0 = time.perf_counter()
    for _ in range(REPEATS):
        expr.to_xml()
    print(f"{(time.perf_counter() - t0) * 1000:.4f}")


if __name__ == "__main__":
    main()
