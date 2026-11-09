"""Time ClientObject.set_property by hydrating one Graph users page.

Hydrate is the offline mapping path: JsonProfile walks the payload and
set_property stores every field. The page is the same shape as a directory
sync (scalars, one datetime, one string collection, etags).
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.directory.users.collection import UserCollection
from office365.graph_client import GraphClient
from office365.runtime.odata.v4.json_format import V4JsonFormat

REPEATS = 20
PAGE_SIZE = 999


def _user_page(n: int) -> dict:
    rows = []
    for i in range(n):
        rows.append(
            {
                "id": f"00000000-0000-0000-0000-{i:012d}",
                "displayName": f"User {i}",
                "givenName": "Ada",
                "surname": "Lovelace",
                "mail": f"user{i}@contoso.com",
                "userPrincipalName": f"user{i}@contoso.com",
                "jobTitle": "Engineer",
                "department": "Research",
                "officeLocation": "London",
                "mobilePhone": "+44-20-7946-0958",
                "accountEnabled": True,
                "usageLocation": "GB",
                "city": "London",
                "country": "United Kingdom",
                "companyName": "Contoso",
                "createdDateTime": "2024-01-15T12:34:56Z",
                "employeeId": f"E-{i:05d}",
                "mailNickname": f"user{i}",
                "preferredLanguage": "en-GB",
                "businessPhones": ["+44-20-7946-0958", "+44-20-7946-0999"],
                "@odata.etag": f'W/"{i}"',
                "@odata.type": "#microsoft.graph.user",
            }
        )
    return {"value": rows}


def _hydrate(client, profile, page):
    collection = UserCollection(client)
    profile.hydrate(collection, page)
    return collection


def main() -> None:
    client = GraphClient()
    profile = V4JsonFormat("minimal")
    page = _user_page(PAGE_SIZE)
    sample = _hydrate(client, profile, page)
    first = sample[0]
    if first.display_name != "User 0" or first.user_principal_name != "user0@contoso.com":
        raise SystemExit("hydrated identity did not match the payload")
    if first.account_enabled is not True:
        raise SystemExit("boolean field was not stored")
    if list(first.business_phones) != ["+44-20-7946-0958", "+44-20-7946-0999"]:
        raise SystemExit("phone collection was not stored")
    if first.created_date_time.year != 2024:
        raise SystemExit("datetime field was not parsed")
    del sample

    t0 = time.perf_counter()
    for _ in range(REPEATS):
        _hydrate(client, profile, page)
    print(f"{(time.perf_counter() - t0) * 1000:.4f}")


if __name__ == "__main__":
    main()
