"""Baseline for office365.runtime.client_value.ClientValue.set_property.

Hydrating a calendar page builds a nested value for every start, end, location,
organizer, attendee and response status. Those values are filled through
set_property. The Event objects are built first; the timed section is the
hydration of one year of a busy shared calendar.
Prints milliseconds for one pass.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from office365.migration.offline import offline_context
from office365.outlook.calendar.events.event import Event
from office365.runtime.odata.v4.json_format import V4JsonFormat

# About a year of meetings on a shared room calendar, two attendees each.
EVENT_COUNT = 1_000


def _payload(index: int) -> dict:
    return {
        "id": f"AAMkAG{index:08d}",
        "subject": f"Quarterly review {index}; budget, headcount",
        "body": {
            "contentType": "html",
            "content": "<p>Agenda: pipeline, headcount, and the March close.</p>",
        },
        "start": {"dateTime": "2026-03-02T09:00:00.0000000", "timeZone": "Europe/Kiev"},
        "end": {"dateTime": "2026-03-02T10:00:00.0000000", "timeZone": "Europe/Kiev"},
        "location": {"displayName": "Room 1, 5th floor"},
        "organizer": {"emailAddress": {"name": "Alice O'Brien", "address": "alice@contoso.com"}},
        "attendees": [
            {
                "emailAddress": {"name": "Bob Smith", "address": "bob@contoso.com"},
                "type": "required",
                "status": {"response": "accepted", "time": "2026-02-20T11:00:00Z"},
            },
            {
                "emailAddress": {"name": "Carol Jones", "address": "carol@contoso.com"},
                "type": "optional",
                "status": {"response": "tentativelyAccepted"},
            },
        ],
        "categories": ["Finance", "Q1"],
        "isAllDay": False,
        "isReminderOn": True,
        "reminderMinutesBeforeStart": 15,
        "showAs": "busy",
        "importance": "high",
        "sensitivity": "normal",
        "lastModifiedDateTime": "2026-02-20T11:30:00Z",
        "isOnlineMeeting": True,
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting"},
    }


def main() -> None:
    context = offline_context()
    profile = V4JsonFormat("minimal")
    events = [Event(context) for _ in range(EVENT_COUNT)]
    payloads = [_payload(i) for i in range(EVENT_COUNT)]
    profile.hydrate(events[0], payloads[0], persist_changes=False)
    if events[0].subject != payloads[0]["subject"]:
        raise SystemExit("hydrate did not store the subject")
    if events[0].attendees[0].status.response != "accepted":
        raise SystemExit("hydrate did not store the nested attendee status")
    for event in events:
        event._properties.clear()
        event._changes.clear()

    def ingest() -> None:
        for event, payload in zip(events, payloads, strict=True):
            event._properties.clear()
            event._changes.clear()
            profile.hydrate(event, payload, persist_changes=False)

    ingest()  # warmup
    started = time.perf_counter()
    ingest()
    elapsed_ms = (time.perf_counter() - started) * 1000
    print(f"{elapsed_ms:.4f}")


if __name__ == "__main__":
    main()
