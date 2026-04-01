"""
JARVIS Notes — Windows version.

Apple Notes doesn't exist on Windows. Notes are stored in a local SQLite
database (the same jarvis.db used by memory.py) and provide the same
interface so server.py works without changes.
"""

import asyncio
import logging
import sqlite3
import time
from pathlib import Path
from typing import Optional

log = logging.getLogger("jarvis.notes")

DB_PATH = Path(__file__).parent / "data" / "jarvis.db"


def _get_db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _init():
    conn = _get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS jarvis_notes (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            title    TEXT    DEFAULT '',
            body     TEXT    NOT NULL,
            folder   TEXT    DEFAULT 'JARVIS',
            created_at REAL NOT NULL,
            updated_at REAL
        )
    """)
    conn.commit()
    conn.close()


_init()


# ── Sync helpers (run in thread pool) ────────────────────────────────────────

def _get_recent_sync(count: int) -> list[dict]:
    conn = _get_db()
    rows = conn.execute(
        "SELECT * FROM jarvis_notes ORDER BY created_at DESC LIMIT ?", (count,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def _read_note_sync(title_search: str) -> Optional[dict]:
    conn = _get_db()
    row = conn.execute(
        "SELECT * FROM jarvis_notes WHERE title LIKE ? ORDER BY updated_at DESC LIMIT 1",
        (f"%{title_search}%",),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def _search_sync(query: str, limit: int) -> list[dict]:
    conn = _get_db()
    rows = conn.execute(
        "SELECT * FROM jarvis_notes WHERE title LIKE ? OR body LIKE ? "
        "ORDER BY updated_at DESC LIMIT ?",
        (f"%{query}%", f"%{query}%", limit),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def _create_sync(title: str, body: str, folder: str) -> dict:
    now = time.time()
    conn = _get_db()
    cur = conn.execute(
        "INSERT INTO jarvis_notes (title, body, folder, created_at, updated_at) VALUES (?,?,?,?,?)",
        (title, body, folder, now, now),
    )
    note_id = cur.lastrowid
    conn.commit()
    conn.close()
    log.info(f"Note created: {title!r}")
    return {"success": True, "id": note_id, "title": title}


# ── Async API (matches macOS notes_access.py interface) ──────────────────────

async def get_recent_notes(count: int = 10) -> list[dict]:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _get_recent_sync, count)


async def read_note(title_search: str) -> Optional[dict]:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _read_note_sync, title_search)


async def search_notes_apple(query: str, limit: int = 10) -> list[dict]:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _search_sync, query, limit)


async def create_apple_note(title: str, content: str, folder: str = "JARVIS") -> dict:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _create_sync, title, content, folder)
