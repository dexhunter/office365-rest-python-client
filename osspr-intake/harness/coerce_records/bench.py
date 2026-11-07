"""Median time to normalize one bulk user import through ``coerce_records``.

600 directory-user rows in the shape ``from_csv`` hands over: declared
columns, a ``passwordProfile/...`` dotted header, a ``"; "`` joined
``businessPhones`` collection, plus ``id`` and ``@odata.type`` cells that an
export leaves on the row. The row list is built once, before the timer. Every
cell looks up the column's declared type again, even though the header set
does not change between rows.
Warmup is not included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.directory.users.user import User
from office365.runtime.converters.csv_reader import coerce_records

REPEATS = 21
WARMUP = 8

_RECORDS = [
    {
        "userPrincipalName": f"user{i}@contoso.com",
        "givenName": "Jordan",
        "displayName": f"Jordan {i}",
        "accountEnabled": "True",
        "officeLocation": "Seattle",
        "jobTitle": "Engineer",
        "department": "Retail",
        "city": "Redmond",
        "mobilePhone": "+1-555-0199",
        "companyName": "Contoso",
        "usageLocation": "US",
        "passwordProfile/password": "S3cret!",
        "passwordProfile/forceChangePasswordNextSignIn": "True",
        "businessPhones": "+1-555-0101; +1-555-0102",
        "id": str(i),
        "@odata.type": "#microsoft.graph.user",
    }
    for i in range(600)
]


def _block() -> None:
    coerce_records(User, _RECORDS)


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
