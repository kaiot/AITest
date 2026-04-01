"""
JARVIS Calendar Access — Windows version.

Reads events from Microsoft Outlook via COM automation (win32com).
Falls back to an empty calendar gracefully if Outlook is not installed.
"""

import asyncio
import logging
import time
from datetime import datetime, timedelta
from typing import Optional

log = logging.getLogger("jarvis.calendar")

_events_cache: list[dict] = []
_cache_time: float = 0.0
CACHE_TTL = 300  # 5 minutes


def _get_outlook_ns():
    """Return Outlook MAPI namespace, or None if unavailable."""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook.GetNamespace("MAPI")
    except Exception as e:
        log.warning(f"Outlook not available: {e}")
        return None


def _fetch_events_sync(hours_ahead: int = 24) -> list[dict]:
    """Synchronously fetch Outlook calendar events (runs in thread pool)."""
    try:
        ns = _get_outlook_ns()
        if not ns:
            return []

        calendar_folder = ns.GetDefaultFolder(9)  # olFolderCalendar
        items = calendar_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        now = datetime.now()
        end_time = now + timedelta(hours=hours_ahead)

        # Outlook restriction format
        fmt = "%m/%d/%Y %H:%M %p"
        restriction = (
            f"[Start] >= '{now.strftime(fmt)}' AND [Start] <= '{end_time.strftime(fmt)}'"
        )
        try:
            items = items.Restrict(restriction)
        except Exception:
            pass  # Use unrestricted list if filter fails

        events = []
        for item in items:
            try:
                start = item.Start
                end = item.End
                all_day = bool(item.AllDayEvent)

                # COM date objects have year/month/day attributes
                start_dt = datetime(start.year, start.month, start.day,
                                    start.hour, start.minute)
                end_dt = datetime(end.year, end.month, end.day,
                                  end.hour, end.minute)

                events.append({
                    "title": str(item.Subject or "Untitled"),
                    "start": "All day" if all_day else start_dt.strftime("%I:%M %p").lstrip("0"),
                    "end": "" if all_day else end_dt.strftime("%I:%M %p").lstrip("0"),
                    "location": str(item.Location or ""),
                    "all_day": all_day,
                    "date": start_dt.strftime("%Y-%m-%d"),
                })

                if len(events) >= 25:
                    break
            except Exception as e:
                log.debug(f"Skipping calendar item: {e}")
                continue

        return events

    except Exception as e:
        log.warning(f"Calendar fetch error: {e}")
        return []


async def _fetch_events_async(hours_ahead: int = 24) -> list[dict]:
    loop = asyncio.get_event_loop()
    try:
        return await asyncio.wait_for(
            loop.run_in_executor(None, _fetch_events_sync, hours_ahead),
            timeout=12,
        )
    except asyncio.TimeoutError:
        log.warning("Calendar fetch timed out")
        return []


async def get_todays_events() -> list[dict]:
    """Get today's calendar events."""
    global _events_cache, _cache_time

    if time.time() - _cache_time < CACHE_TTL:
        today = datetime.now().strftime("%Y-%m-%d")
        return [e for e in _events_cache if e.get("date") == today]

    events = await _fetch_events_async(hours_ahead=24)
    _events_cache = events
    _cache_time = time.time()

    today = datetime.now().strftime("%Y-%m-%d")
    return [e for e in events if e.get("date") == today]


async def get_upcoming_events(hours: int = 8) -> list[dict]:
    """Get events starting within the next N hours."""
    return await _fetch_events_async(hours_ahead=hours)


async def get_next_event() -> Optional[dict]:
    """Get the single next upcoming event."""
    events = await get_upcoming_events(hours=24)
    return events[0] if events else None


def format_events_for_context(events: list[dict]) -> str:
    """Format events for injection into the LLM system prompt."""
    if not events:
        return "No upcoming events."
    lines = []
    for e in events:
        time_str = e.get("start", "")
        title = e.get("title", "Untitled")
        loc = f" @ {e['location']}" if e.get("location") else ""
        lines.append(f"  {time_str}: {title}{loc}")
    return "Calendar:\n" + "\n".join(lines)


def format_schedule_summary(events: list[dict]) -> str:
    """Format a brief spoken schedule summary."""
    if not events:
        return "Your calendar is clear, sir."
    count = len(events)
    if count == 1:
        e = events[0]
        return f"One event today: {e['title']} at {e['start']}."
    nxt = events[0]
    return f"{count} events today. Next up: {nxt['title']} at {nxt['start']}."


async def refresh_cache():
    """Force-refresh the calendar cache."""
    global _events_cache, _cache_time
    _events_cache = await _fetch_events_async(hours_ahead=24)
    _cache_time = time.time()
