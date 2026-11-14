"""Baseline for office365.runtime.client_object.ClientObject.get_property.

A directory export reads the columns it selected on every loaded user. The
values are already in the property bag (the page was hydrated). Setup,
including hydration, is excluded from the timed section.
Prints milliseconds for one pass.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.directory.users.user import User
from office365.migration.offline import offline_context
from office365.runtime.odata.v4.json_format import V4JsonFormat

# A mid-size tenant directory, one full export rather than a single Graph page.
USER_COUNT = 18_000
FIELDS = (
    "id",
    "displayName",
    "givenName",
    "surname",
    "mail",
    "userPrincipalName",
    "jobTitle",
    "department",
    "officeLocation",
    "accountEnabled",
    "createdDateTime",
    "city",
    "country",
    "companyName",
    "usageLocation",
    "preferredLanguage",
    "employeeId",
)


def main() -> None:
    context = offline_context()
    profile = V4JsonFormat("minimal")
    users = []
    for i in range(USER_COUNT):
        user = User(context)
        profile.hydrate(
            user,
            {
                "id": f"00000000-0000-0000-0000-{i:012d}",
                "displayName": f"Alex Morgan {i}",
                "givenName": "Alex",
                "surname": "Morgan",
                "mail": f"alex.morgan{i}@contoso.com",
                "userPrincipalName": f"alex.morgan{i}@contoso.com",
                "jobTitle": "Finance Analyst",
                "department": "Finance" if i % 2 else "Engineering",
                "officeLocation": "Building 5, London",
                "accountEnabled": True,
                "createdDateTime": "2024-03-02T09:14:22Z",
                "city": "London",
                "country": "United Kingdom",
                "companyName": "Contoso Ltd",
                "usageLocation": "GB",
                "preferredLanguage": "en-GB",
                "employeeId": f"E{i:05d}",
            },
            persist_changes=False,
        )
        users.append(user)
    if users[0].get_property("displayName") != "Alex Morgan 0":
        raise SystemExit("get_property did not return the hydrated display name")
    users[0].get_property("mail")  # warmup
    started = time.perf_counter()
    for user in users:
        for name in FIELDS:
            user.get_property(name)
    elapsed_ms = (time.perf_counter() - started) * 1000
    print(f"{elapsed_ms:.4f}")


if __name__ == "__main__":
    main()
