"""
JARVIS Work Mode — Claude Code session manager (Windows version).

Manages persistent Claude Code CLI sessions per project directory.
Uses claude -p with --continue for multi-turn project context.
"""

import asyncio
import json
import logging
import re
import time
from pathlib import Path
from typing import Optional

log = logging.getLogger("jarvis.work_mode")

SESSIONS_FILE = Path(__file__).parent / "data" / "work_sessions.json"
SESSION_TIMEOUT = 300  # 5 minutes inactivity

# Patterns that indicate casual conversation (no Claude Code needed)
_CASUAL_PATTERNS = [
    r"^(hi|hello|hey|what's up|sup)\b",
    r"^(what time|what day|what('s| is) the (date|weather|time))\b",
    r"^(thanks?|thank you|thx|ok|okay|got it|sounds good|great|perfect|awesome|cool)\b",
    r"^(yes|no|sure|alright|yep|nope|yup)\b",
    r"^(tell me (a joke|something funny))\b",
    r"^(who are you|what are you|what can you do|how are you)\b",
    r"^(good (morning|afternoon|evening|night))\b",
]
_casual_re = [re.compile(p, re.IGNORECASE) for p in _CASUAL_PATTERNS]


def is_casual_question(text: str) -> bool:
    """Return True if this is casual chat that doesn't need Claude Code."""
    text = text.strip()
    for pattern in _casual_re:
        if pattern.match(text):
            return True
    # Short phrases (< 6 words) without technical keywords are usually casual
    words = text.split()
    if len(words) <= 5:
        tech_words = {
            "build", "create", "fix", "debug", "code", "deploy", "test",
            "implement", "refactor", "add", "update", "change", "delete",
            "error", "bug", "feature", "function", "class", "file", "git",
        }
        if not any(w.lower() in tech_words for w in words):
            return True
    return False


class WorkSession:
    """Manages a Claude Code CLI session scoped to a project directory."""

    def __init__(self):
        self._project_dir: Optional[str] = None
        self._active: bool = False
        self._last_used: float = 0.0
        self._message_count: int = 0
        self._sessions: dict = {}
        self._load()

    def _load(self):
        try:
            SESSIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
            if SESSIONS_FILE.exists():
                self._sessions = json.loads(SESSIONS_FILE.read_text())
        except Exception:
            self._sessions = {}

    def _save(self):
        try:
            SESSIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
            SESSIONS_FILE.write_text(json.dumps(self._sessions, indent=2))
        except Exception as e:
            log.debug(f"Could not save sessions: {e}")

    async def start(self, project_dir: str) -> bool:
        """Begin a work session for the given project directory."""
        self._project_dir = project_dir
        self._active = True
        self._last_used = time.time()
        session_data = self._sessions.get(project_dir, {})
        self._message_count = session_data.get("message_count", 0)
        log.info(f"Work session started: {project_dir} (msg #{self._message_count})")
        return True

    async def stop(self):
        """End the current work session and persist state."""
        if self._project_dir:
            self._sessions[self._project_dir] = {
                "message_count": self._message_count,
                "last_used": self._last_used,
            }
            self._save()
        self._active = False
        self._project_dir = None
        log.info("Work session ended")

    async def send(self, prompt: str, timeout: int = 300) -> str:
        """Send a prompt to Claude Code in the project directory and return output."""
        if not self._active or not self._project_dir:
            return "No active work session, sir."

        self._last_used = time.time()
        self._message_count += 1

        project_path = Path(self._project_dir)
        use_continue = self._message_count > 1

        cmd = ["claude", "-p"]
        if use_continue:
            cmd.append("--continue")
        cmd += ["--dangerously-skip-permissions", prompt]

        try:
            proc = await asyncio.create_subprocess_exec(
                *cmd,
                cwd=str(project_path),
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                # On Windows we don't need creationflags for background
            )

            try:
                stdout, stderr = await asyncio.wait_for(proc.communicate(), timeout=timeout)
            except asyncio.TimeoutError:
                proc.kill()
                return f"Task timed out after {timeout}s, sir."

            output = stdout.decode(errors="replace").strip()
            if not output:
                output = stderr.decode(errors="replace").strip()

            # Persist state
            self._sessions[str(project_path)] = {
                "message_count": self._message_count,
                "last_used": self._last_used,
            }
            self._save()

            return output or "Task completed, sir."

        except FileNotFoundError:
            return (
                "Claude Code is not installed or not in PATH, sir. "
                "Install it with: npm install -g @anthropic-ai/claude-code"
            )
        except Exception as e:
            log.error(f"WorkSession.send failed: {e}")
            return f"Error running Claude Code: {e}"

    @property
    def is_active(self) -> bool:
        if not self._active:
            return False
        if time.time() - self._last_used > SESSION_TIMEOUT:
            self._active = False
            return False
        return True

    @property
    def project_dir(self) -> Optional[str]:
        return self._project_dir
