"""
JARVIS Mail Access — Windows Outlook version.

Read-only access to Microsoft Outlook via COM automation.
Intentionally no send/delete/move — read only.
"""

import asyncio
import logging
from typing import Optional

log = logging.getLogger("jarvis.mail")


def _get_ns():
    """Return Outlook MAPI namespace or None."""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook.GetNamespace("MAPI")
    except Exception as e:
        log.warning(f"Outlook not available: {e}")
        return None


def _accounts_sync() -> list[dict]:
    try:
        ns = _get_ns()
        if not ns:
            return []
        return [
            {"name": str(acc.DisplayName or ""), "email": str(acc.SmtpAddress or "")}
            for acc in ns.Accounts
        ]
    except Exception as e:
        log.warning(f"accounts error: {e}")
        return []


def _unread_sync() -> dict:
    try:
        ns = _get_ns()
        if not ns:
            return {}
        inbox = ns.GetDefaultFolder(6)  # olFolderInbox
        return {"Inbox": inbox.UnReadItemCount}
    except Exception as e:
        log.warning(f"unread count error: {e}")
        return {}


def _messages_sync(limit: int = 10, unread_only: bool = False) -> list[dict]:
    try:
        ns = _get_ns()
        if not ns:
            return []

        inbox = ns.GetDefaultFolder(6)
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # newest first

        results = []
        for item in items:
            if len(results) >= limit:
                break
            try:
                if unread_only and not item.UnRead:
                    continue
                received = item.ReceivedTime
                received_str = (
                    f"{received.year}-{received.month:02d}-{received.day:02d} "
                    f"{received.hour:02d}:{received.minute:02d}"
                )
                results.append({
                    "id": str(item.EntryID),
                    "subject": str(item.Subject or ""),
                    "sender": str(item.SenderName or ""),
                    "sender_email": str(getattr(item, "SenderEmailAddress", "") or ""),
                    "received": received_str,
                    "unread": bool(item.UnRead),
                    "preview": str(item.Body or "")[:300].replace("\r\n", " ").replace("\n", " "),
                })
            except Exception as e:
                log.debug(f"Skipping mail item: {e}")
                continue
        return results
    except Exception as e:
        log.warning(f"messages fetch error: {e}")
        return []


def _search_sync(query: str, limit: int = 10) -> list[dict]:
    try:
        ns = _get_ns()
        if not ns:
            return []
        inbox = ns.GetDefaultFolder(6)
        items = inbox.Items
        try:
            q = query.replace("'", "''")
            items = items.Restrict(
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{q}%' "
                f"OR \"urn:schemas:httpmail:fromname\" LIKE '%{q}%'"
            )
        except Exception:
            pass

        results = []
        for item in items:
            if len(results) >= limit:
                break
            try:
                received = item.ReceivedTime
                results.append({
                    "id": str(item.EntryID),
                    "subject": str(item.Subject or ""),
                    "sender": str(item.SenderName or ""),
                    "received": str(received),
                    "preview": str(item.Body or "")[:200],
                })
            except Exception:
                continue
        return results
    except Exception as e:
        log.warning(f"mail search error: {e}")
        return []


def _read_sync(entry_id: str) -> str:
    try:
        ns = _get_ns()
        if not ns:
            return "Unable to access Outlook."
        item = ns.GetItemFromID(entry_id)
        return str(item.Body or "")
    except Exception as e:
        return f"Could not read message: {e}"


# ── Async wrappers ────────────────────────────────────────────────────────────

async def _run(fn, *args, timeout: int = 10):
    loop = asyncio.get_event_loop()
    try:
        return await asyncio.wait_for(loop.run_in_executor(None, fn, *args), timeout)
    except asyncio.TimeoutError:
        return None


async def get_accounts() -> list[dict]:
    return (await _run(_accounts_sync)) or []


async def get_unread_count() -> dict:
    return (await _run(_unread_sync)) or {}


async def get_unread_messages(limit: int = 10) -> list[dict]:
    loop = asyncio.get_event_loop()
    try:
        return await asyncio.wait_for(
            loop.run_in_executor(None, _messages_sync, limit, True), 10
        ) or []
    except asyncio.TimeoutError:
        return []


async def get_recent_messages(limit: int = 10) -> list[dict]:
    loop = asyncio.get_event_loop()
    try:
        return await asyncio.wait_for(
            loop.run_in_executor(None, _messages_sync, limit, False), 10
        ) or []
    except asyncio.TimeoutError:
        return []


async def search_mail(query: str, limit: int = 10) -> list[dict]:
    loop = asyncio.get_event_loop()
    try:
        return await asyncio.wait_for(
            loop.run_in_executor(None, _search_sync, query, limit), 10
        ) or []
    except asyncio.TimeoutError:
        return []


async def read_message(message_id: str) -> str:
    return (await _run(_read_sync, message_id)) or "Could not read message."


# ── Formatters ────────────────────────────────────────────────────────────────

def format_unread_summary(counts: Optional[dict] = None) -> str:
    if not counts:
        return "No unread mail, sir."
    total = sum(counts.values())
    if total == 0:
        return "No unread mail, sir."
    return f"You have {total} unread message{'s' if total > 1 else ''}, sir."


def format_messages_for_context(messages: list[dict]) -> str:
    if not messages:
        return "No recent messages."
    lines = ["Recent email:"]
    for m in messages[:5]:
        flag = " [UNREAD]" if m.get("unread") else ""
        lines.append(f"  From: {m['sender']}{flag} — {m['subject']}")
    return "\n".join(lines)


def format_messages_for_voice(messages: list[dict]) -> str:
    if not messages:
        return "No messages to read, sir."
    count = len(messages)
    result = f"{count} message{'s' if count > 1 else ''}. "
    for m in messages[:3]:
        result += f"From {m['sender']}: {m['subject']}. "
    if count > 3:
        result += f"And {count - 3} more."
    return result
