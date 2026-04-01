"""
Patches server.py to replace macOS-specific code with Windows equivalents.
Run once after downloading the original server.py from GitHub.
"""

import re
import sys
from pathlib import Path

SERVER = Path(__file__).parent / "server.py"


def patch(content: str) -> str:
    # ── 1. Replace AppleScript terminal spawning in _run_task ────────────────
    old_terminal = r'''        applescript = f\'\'\'
        tell application "Terminal"
            activate
            set newTab to do script "cd {work_dir} && cat .jarvis_prompt.md | claude -p --dangerously-skip-permissions | tee .jarvis_output.txt; echo \'\\\\n--- JARVIS TASK COMPLETE ---\'"
        end tell
        \'\'\'

        process = await asyncio.create_subprocess_exec(
            "osascript", "-e", applescript,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        await process.communicate()
        task.pid = process.pid'''

    new_terminal = '''        # Windows: write a batch runner, open Windows Terminal, monitor output
        import subprocess as _sp
        bat_file = Path(work_dir) / ".jarvis_run.bat"
        out_file = Path(work_dir) / ".jarvis_output.txt"
        bat_file.write_text(
            f\'@echo off\\n\'
            f\'type ".jarvis_prompt.md" | claude -p --dangerously-skip-permissions > ".jarvis_output.txt" 2>&1\\n\'
            f\'echo --- JARVIS TASK COMPLETE --- >> ".jarvis_output.txt"\\n\',
            encoding="utf-8",
        )
        try:
            process = await asyncio.create_subprocess_exec(
                "wt", "-d", work_dir, "cmd", "/k", str(bat_file),
                stdout=asyncio.subprocess.DEVNULL,
                stderr=asyncio.subprocess.DEVNULL,
            )
        except FileNotFoundError:
            _sp.Popen(f\'start cmd /k "{bat_file}"\', shell=True, cwd=work_dir)
            process = None
        task.pid = process.pid if process else None'''

    content = content.replace(old_terminal, new_terminal)

    # ── 2. Replace _focus_terminal_window ────────────────────────────────────
    old_focus = r'''async def _focus_terminal_window(project_name: str):
    """Bring a Terminal window matching the project name to front."""
    escaped = project_name.replace('"', '\\"')
    script = f\'\'\'
tell application "Terminal"
    repeat with w in windows
        if name of w contains "{escaped}" then
            set index of w to 1
            activate
            exit repeat
        end if
    end repeat
end tell
\'\'\'
    try:
        proc = await asyncio.create_subprocess_exec(
            "osascript", "-e", script,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        await asyncio.wait_for(proc.communicate(), timeout=5)
    except Exception:
        pass'''

    new_focus = '''async def _focus_terminal_window(project_name: str):
    """Bring a Windows Terminal window matching the project name to front."""
    try:
        import win32gui
        import win32con

        def _enum(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if project_name.lower() in title.lower():
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                    except Exception:
                        pass
                    return False
            return True

        win32gui.EnumWindows(_enum, None)
    except Exception:
        pass'''

    content = content.replace(old_focus, new_focus)

    # ── 3. Replace api_fix_self AppleScript ──────────────────────────────────
    old_fix = r'''    jarvis_dir = str(Path(__file__).parent)
    # The work_session is per-WebSocket, so we set a flag that the handler picks up
    # For now, also open Terminal so user can see
    script = (
        \'tell application "Terminal"\\n\'
        \'    activate\\n\'
        f\'    do script "cd {jarvis_dir} && claude --dangerously-skip-permissions"\\n\'
        \'end tell\'
    )
    await asyncio.create_subprocess_exec(
        "osascript", "-e", script,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )'''

    new_fix = '''    jarvis_dir = str(Path(__file__).parent)
    import subprocess as _sp
    try:
        await asyncio.create_subprocess_exec(
            "wt", "-d", jarvis_dir, "cmd", "/k", "claude --dangerously-skip-permissions",
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.DEVNULL,
        )
    except FileNotFoundError:
        _sp.Popen(
            f\'start cmd /k "cd /d "{jarvis_dir}" && claude --dangerously-skip-permissions"\',
            shell=True,
        )'''

    content = content.replace(old_fix, new_fix)

    # ── 4. Update system prompt: macOS → Windows references ──────────────────
    replacements = [
        ("Terminal.app", "Windows Terminal"),
        ("You CAN open Terminal.app via AppleScript",
         "You CAN open Windows Terminal"),
        ("Google Chrome and browse",
         "Chrome browser and browse"),
        ("open windows, active apps, and screenshot vision",
         "open windows, active apps, and screenshot vision (Windows)"),
        ("You CAN read {user_name}'s calendar — today's events, upcoming meetings, schedule overview",
         "You CAN read {user_name}'s Outlook calendar — today's events, upcoming meetings, schedule overview"),
        ("You CAN read {user_name}'s email (READ-ONLY) — unread count, recent messages, search by sender/subject. You CANNOT send, delete, or modify emails.",
         "You CAN read {user_name}'s Outlook email (READ-ONLY) — unread count, recent messages, search by sender/subject. You CANNOT send, delete, or modify emails."),
        ("You CAN read Apple Notes and create NEW notes",
         "You CAN read JARVIS notes and create NEW notes (stored locally)"),
        ("osascript", "# osascript not available on Windows"),
    ]
    for old, new in replacements:
        content = content.replace(old, new)

    return content


def main():
    if not SERVER.exists():
        print(f"ERROR: {SERVER} not found. Run setup.bat first.")
        sys.exit(1)

    original = SERVER.read_text(encoding="utf-8")
    patched = patch(original)

    if patched == original:
        print("WARNING: No patches applied — server.py may already be patched or source changed.")
    else:
        SERVER.write_text(patched, encoding="utf-8")
        print("server.py patched for Windows successfully.")


if __name__ == "__main__":
    main()
