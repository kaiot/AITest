"""
JARVIS Action Executor — Windows system actions.

Execute actions IMMEDIATELY, before generating any LLM response.
Each function returns {"success": bool, "confirmation": str}.
"""

import asyncio
import logging
import os
import re
import subprocess
import time
import webbrowser
from pathlib import Path
from urllib.parse import quote

log = logging.getLogger("jarvis.actions")

DESKTOP_PATH = Path.home() / "Desktop"


async def _open_wt(cwd: str = None, command: str = None) -> bool:
    """Try to open Windows Terminal. Returns True on success."""
    try:
        args = ["wt"]
        if cwd:
            args += ["-d", cwd]
        if command:
            args += ["cmd", "/k", command]
        proc = await asyncio.create_subprocess_exec(
            *args,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.DEVNULL,
        )
        await asyncio.sleep(0.5)
        return True
    except FileNotFoundError:
        return False


async def open_terminal(command: str = "") -> dict:
    """Open Windows Terminal and optionally run a command."""
    try:
        if command:
            success = await _open_wt(command=command)
            if not success:
                subprocess.Popen(
                    f'start cmd /k "{command}"',
                    shell=True,
                )
        else:
            success = await _open_wt()
            if not success:
                subprocess.Popen("start cmd", shell=True)
        return {"success": True, "confirmation": "Terminal is open, sir."}
    except Exception as e:
        log.error(f"open_terminal failed: {e}")
        return {"success": False, "confirmation": "Had trouble opening the terminal, sir."}


async def open_browser(url: str, browser: str = "chrome") -> dict:
    """Open URL in the user's browser."""
    try:
        if browser.lower() == "firefox":
            try:
                subprocess.Popen(["firefox", url])
                return {"success": True, "confirmation": "Pulled that up in Firefox, sir."}
            except FileNotFoundError:
                pass

        # Try Chrome explicitly, then fall back to system default
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        for chrome in chrome_paths:
            if Path(chrome).exists():
                subprocess.Popen([chrome, url])
                return {"success": True, "confirmation": "Pulled that up in Chrome, sir."}

        # System default browser
        os.startfile(url)
        return {"success": True, "confirmation": "Pulled that up in your browser, sir."}

    except Exception as e:
        log.error(f"open_browser failed: {e}")
        # Last resort
        try:
            webbrowser.open(url)
            return {"success": True, "confirmation": "Pulled that up in your browser, sir."}
        except Exception:
            return {"success": False, "confirmation": "Had trouble opening the browser, sir."}


async def open_chrome(url: str) -> dict:
    return await open_browser(url, "chrome")


async def open_claude_in_project(project_dir: str, prompt: str) -> dict:
    """Open Windows Terminal, cd to project dir, run Claude Code interactively."""
    claude_md = Path(project_dir) / "CLAUDE.md"
    claude_md.write_text(
        f"# Task\n\n{prompt}\n\nBuild this completely. "
        "If web app, make index.html work standalone.\n"
    )

    try:
        success = await _open_wt(
            cwd=project_dir,
            command="claude --dangerously-skip-permissions",
        )
        if not success:
            bat = Path(project_dir) / ".jarvis_start.bat"
            bat.write_text(
                f'@echo off\ncd /d "{project_dir}"\nclaude --dangerously-skip-permissions\n'
            )
            subprocess.Popen(f'start cmd /k "{bat}"', shell=True)

        return {
            "success": True,
            "confirmation": "Claude Code is running in Terminal, sir. You can watch the progress.",
        }
    except Exception as e:
        log.error(f"open_claude_in_project failed: {e}")
        return {"success": False, "confirmation": "Had trouble spawning Claude Code, sir."}


async def prompt_existing_terminal(project_name: str, prompt: str) -> dict:
    """Find a terminal window matching a project name and type a prompt into it."""
    try:
        import win32gui
        import win32con

        target_hwnd = None

        def _enum(hwnd, _):
            nonlocal target_hwnd
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if project_name.lower() in title.lower():
                    target_hwnd = hwnd
            return True

        win32gui.EnumWindows(_enum, None)

        if not target_hwnd:
            return {
                "success": False,
                "confirmation": f"Couldn't find a terminal for {project_name}, sir.",
            }

        win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(target_hwnd)
        await asyncio.sleep(0.5)

        try:
            from pynput.keyboard import Controller, Key
            kb = Controller()
            kb.type(prompt)
            kb.press(Key.enter)
            kb.release(Key.enter)
        except ImportError:
            import pyautogui
            pyautogui.typewrite(prompt, interval=0.02)
            pyautogui.press("enter")

        return {"success": True, "confirmation": f"Sent that to {project_name}, sir."}

    except ImportError:
        return {
            "success": False,
            "confirmation": "Terminal automation requires pywin32, sir.",
        }
    except asyncio.TimeoutError:
        return {"success": False, "confirmation": "Terminal operation timed out, sir."}
    except Exception as e:
        log.error(f"prompt_existing_terminal failed: {e}")
        return {"success": False, "confirmation": "Something went wrong reaching that terminal, sir."}


async def get_chrome_tab_info() -> dict:
    """Read the current Chrome tab's title and URL via Chrome debugging protocol."""
    try:
        import httpx
        async with httpx.AsyncClient(timeout=3.0) as client:
            resp = await client.get("http://localhost:9222/json")
            if resp.status_code == 200:
                tabs = resp.json()
                for tab in tabs:
                    if tab.get("type") == "page":
                        url = tab.get("url", "")
                        if url.startswith("http") and "chrome://" not in url:
                            return {"title": tab.get("title", ""), "url": url}
    except Exception as e:
        log.debug(f"get_chrome_tab_info failed: {e}")
    return {}


async def monitor_build(project_dir: str, ws=None, synthesize_fn=None) -> None:
    """Monitor a Claude Code build for completion. Notify via WebSocket when done."""
    import base64

    output_file = Path(project_dir) / ".jarvis_output.txt"
    start = time.time()
    timeout = 600  # 10 minutes

    while time.time() - start < timeout:
        await asyncio.sleep(5)
        if output_file.exists():
            content = output_file.read_text(errors="replace")
            if "--- JARVIS TASK COMPLETE ---" in content:
                log.info(f"Build complete in {project_dir}")
                if ws and synthesize_fn:
                    try:
                        msg = "The build is complete, sir."
                        audio_bytes = await synthesize_fn(msg)
                        if audio_bytes:
                            encoded = base64.b64encode(audio_bytes).decode()
                            await ws.send_json({"type": "status", "state": "speaking"})
                            await ws.send_json({"type": "audio", "data": encoded, "text": msg})
                            await ws.send_json({"type": "status", "state": "idle"})
                    except Exception as e:
                        log.warning(f"Build notification failed: {e}")
                return

    log.warning(f"Build timed out in {project_dir}")


async def execute_action(intent: dict, projects: list = None) -> dict:
    """Route a classified intent to the right action function."""
    action = intent.get("action", "chat")
    target = intent.get("target", "")

    if action == "open_terminal":
        result = await open_terminal("claude --dangerously-skip-permissions")
        result["project_dir"] = None
        return result

    elif action == "browse":
        if target.startswith("http://") or target.startswith("https://"):
            url = target
        else:
            url = f"https://www.google.com/search?q={quote(target)}"

        browser = "firefox" if "firefox" in target.lower() else "chrome"
        result = await open_browser(url, browser)
        result["project_dir"] = None
        return result

    elif action == "build":
        project_name = _generate_project_name(target)
        project_dir = str(DESKTOP_PATH / project_name)
        os.makedirs(project_dir, exist_ok=True)
        result = await open_claude_in_project(project_dir, target)
        result["project_dir"] = project_dir
        return result

    else:
        return {"success": False, "confirmation": "", "project_dir": None}


def _generate_project_name(prompt: str) -> str:
    """Generate a kebab-case project folder name from the prompt."""
    quoted = re.search(r'"([^"]+)"', prompt)
    if quoted:
        name = quoted.group(1).strip()
        name = re.sub(r"[^a-zA-Z0-9\s-]", "", name).strip()
        if name:
            return re.sub(r"[\s]+", "-", name.lower())

    called = re.search(r'(?:called|named)\s+(\S+(?:[-_]\S+)*)', prompt, re.IGNORECASE)
    if called:
        name = re.sub(r"[^a-zA-Z0-9-]", "", called.group(1))
        if len(name) > 3:
            return name.lower()

    words = re.sub(r"[^a-zA-Z0-9\s]", "", prompt.lower()).split()
    skip = {
        "a", "the", "an", "me", "build", "create", "make", "for", "with", "and",
        "to", "of", "i", "want", "need", "new", "project", "directory", "called",
        "on", "desktop", "that", "application", "app", "full", "stack", "simple",
        "web", "page", "site", "named",
    }
    meaningful = [w for w in words if w not in skip and len(w) > 2][:4]
    return "-".join(meaningful) if meaningful else "jarvis-project"
