"""Median time to read the properties named by a batch of OData filters.

250 filters, each the kind a tenant scan puts on a user, message, or list
query: string literals, a datetime literal, function calls, and ``and``/``or``.
The expressions are built once, before the timer. ``filter`` / ``order_by``
ask ``referenced_properties`` which names the expression touches.
The scan rewrites the string twice with regular expressions, then walks every
identifier and slices the tail to see whether it is a function call.
Warmup is not included. ``gc.collect`` runs between repeats, outside the timer.
"""

from __future__ import annotations

import gc
import statistics
import time

from office365.runtime.query_capabilities import referenced_properties

REPEATS = 21
WARMUP = 8

_FILTERS = []
for i in range(250):
    _FILTERS.append(
        " and ".join(
            (
                f"startswith(displayName,'User {i}')",
                "accountEnabled eq true",
                f"mail eq 'user{i}@contoso.com'",
                "createdDateTime ge 2024-01-15T12:34:56Z",
                f"(department eq 'Retail' or officeLocation eq 'Building {i % 6}')",
                "not(userType eq 'Guest')",
            )
        )
    )


def _block() -> None:
    for expression in _FILTERS:
        referenced_properties(expression)


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
