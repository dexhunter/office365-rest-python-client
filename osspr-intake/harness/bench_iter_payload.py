"""Time JsonProfile.iter_payload on one Graph users page.

A directory sync reads a page of users (the default Graph page size) with the
control information a minimal-metadata response actually carries: an etag and
a type on every entity, plus the page's next link. The walk is offline.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.runtime.odata.v4.json_format import V4JsonFormat

REPEATS = 160
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
    return {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users",
        "value": rows,
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=page-2",
    }


def main() -> None:
    page = _user_page(PAGE_SIZE)
    profile = V4JsonFormat("minimal")
    walked = list(profile.iter_payload(page))
    if walked[0] != ("__nextLinkUrl", page["@odata.nextLink"]):
        raise SystemExit("next link was not surfaced")
    if len(walked) != PAGE_SIZE + 1:
        raise SystemExit(f"expected {PAGE_SIZE + 1} entries, got {len(walked)}")
    entity = walked[1][1]
    if "@odata.etag" in entity or entity.get("__etag") != 'W/"0"':
        raise SystemExit("etag was not rewritten to __etag")
    if entity["displayName"] != "User 0" or entity["businessPhones"] != page["value"][0]["businessPhones"]:
        raise SystemExit("entity fields did not round-trip")

    t0 = time.perf_counter()
    for _ in range(REPEATS):
        list(profile.iter_payload(page))
    print(f"{(time.perf_counter() - t0) * 1000:.4f}")


if __name__ == "__main__":
    main()
