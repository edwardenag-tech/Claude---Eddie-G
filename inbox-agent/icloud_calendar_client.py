"""Apple iCloud calendar wrapper using CalDAV.

Authenticates with an app-specific password (never the main Apple ID
password) against Apple's CalDAV server. Read-only: this client only
fetches events, it never writes to the calendar.
"""

import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import caldav
import pytz

logger = logging.getLogger(__name__)

ICLOUD_CALDAV_URL = "https://caldav.icloud.com"

# Network calls must never hang an unattended overnight run -- if Apple's
# CalDAV server is slow or unreachable, fail fast and let the caller degrade
# gracefully rather than blocking the whole agent.
_REQUEST_TIMEOUT_SECS = 15


class ICloudCalendarClient:
    def __init__(self, apple_id: str, app_password: str, timeout: int = _REQUEST_TIMEOUT_SECS):
        """Connect to iCloud's CalDAV server and discover the account's calendars.

        Raises on failure (bad credentials, network error, timeout) --
        callers should catch this and degrade gracefully, same as
        build_outlook_client in agent.py.
        """
        self.apple_id = apple_id
        self._client = caldav.DAVClient(
            url=ICLOUD_CALDAV_URL,
            username=apple_id,
            password=app_password,
            timeout=timeout,
        )
        principal = self._client.principal()
        self._calendars = principal.calendars()
        logger.info("iCloud CalDAV connected -- %d calendar(s) found", len(self._calendars))

    def get_todays_events(self, tz_name: str = "Australia/Sydney") -> List[Dict]:
        """Today's events (local calendar day) across all personal calendars, sorted by start time."""
        tz = pytz.timezone(tz_name)
        start_local = datetime.now(tz).replace(hour=0, minute=0, second=0, microsecond=0)
        end_local = start_local + timedelta(days=1)

        events: List[Dict] = []
        for calendar in self._calendars:
            try:
                results = calendar.search(
                    start=start_local, end=end_local, event=True, expand=True,
                )
            except Exception as exc:
                logger.warning("iCloud calendar '%s' search failed (continuing): %s", getattr(calendar, "name", "?"), exc)
                continue

            for result in results:
                try:
                    events.append(self._extract_event_data(result, tz))
                except Exception as exc:
                    logger.warning("Failed to parse an iCloud event (skipping): %s", exc)

        events.sort(key=lambda e: e["start"])
        return events

    @staticmethod
    def _extract_event_data(caldav_event, tz) -> Dict:
        """Flatten a caldav Event's VEVENT component into the same shape used
        by OutlookClient._extract_event_data (subject/start/end/location/organizer/is_all_day)."""
        vevent = caldav_event.icalendar_component

        dtstart = vevent["dtstart"].dt
        dtend_prop = vevent.get("dtend")
        dtend = dtend_prop.dt if dtend_prop is not None else dtstart

        is_all_day = not isinstance(dtstart, datetime)

        def _fmt(d):
            if isinstance(d, datetime):
                if d.tzinfo is not None:
                    d = d.astimezone(tz)
                return d.strftime("%Y-%m-%dT%H:%M:%S")
            return d.strftime("%Y-%m-%dT00:00:00")

        organizer = vevent.get("organizer")
        organizer_str = ""
        if organizer:
            organizer_str = str(organizer).replace("mailto:", "")
            cn = organizer.params.get("CN") if hasattr(organizer, "params") else None
            if cn:
                organizer_str = cn

        return {
            "subject": str(vevent.get("summary", "(no subject)")),
            "start": _fmt(dtstart),
            "end": _fmt(dtend),
            "location": str(vevent.get("location", "")),
            "organizer": organizer_str,
            "is_all_day": is_all_day,
        }
