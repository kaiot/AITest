"""
JARVIS Screen Awareness — Windows version.

Two capabilities:
1. Window/app list via pywin32 + psutil (fast, text-based)
2. Screenshot via mss or Pillow -> Claude vision API (sees everything)
"""

import asyncio
import base64
import logging
import tempfile
from pathlib import Path

log = logging.getLogger("jarvis.screen")


async def get_active_windows() -> list[dict]:
    """Get list of visible windows with app name, window title, and focus state."""
    try:
        import win32gui
        import win32process
        import psutil

        windows = []
        foreground_hwnd = win32gui.GetForegroundWindow()

        _SKIP_APPS = {
            "explorer", "searchapp", "startmenuexperiencehost",
            "shellexperiencehost", "textinputhost", "applicationframehost",
            "systemsettings", "lockapp",
        }

        def _enum(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title or len(title) < 2:
                return True
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                proc = psutil.Process(pid)
                app_name = proc.name().replace(".exe", "")
            except Exception:
                app_name = "Unknown"

            if app_name.lower() in _SKIP_APPS:
                return True

            windows.append({
                "app": app_name,
                "title": title,
                "frontmost": hwnd == foreground_hwnd,
            })
            return True

        win32gui.EnumWindows(_enum, None)
        return windows

    except ImportError:
        # Fallback: PowerShell window list (no titles, just process names)
        try:
            proc = await asyncio.create_subprocess_exec(
                "powershell", "-NoProfile", "-Command",
                "Get-Process | Where-Object {$_.MainWindowTitle} | "
                "Select-Object ProcessName,MainWindowTitle | ConvertTo-Csv -NoTypeInformation",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )
            stdout, _ = await asyncio.wait_for(proc.communicate(), timeout=8)
            windows = []
            for line in stdout.decode(errors="replace").splitlines()[1:]:
                parts = line.strip('"').split('","')
                if len(parts) >= 2:
                    windows.append({
                        "app": parts[0].replace(".exe", ""),
                        "title": parts[1],
                        "frontmost": False,
                    })
            return windows
        except Exception as e:
            log.warning(f"Fallback window list failed: {e}")
            return []
    except Exception as e:
        log.warning(f"get_active_windows error: {e}")
        return []


async def get_running_apps() -> list[str]:
    """Get list of running application names (visible processes only)."""
    try:
        import psutil
        apps = set()
        for proc in psutil.process_iter(["name", "status"]):
            try:
                if proc.info["status"] == psutil.STATUS_RUNNING:
                    name = proc.info["name"].replace(".exe", "")
                    _skip = {"svchost", "RuntimeBroker", "conhost", "System", "lsass",
                              "csrss", "smss", "wininit", "services", "dwm"}
                    if name and name not in _skip:
                        apps.add(name)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return sorted(apps)
    except ImportError:
        try:
            proc = await asyncio.create_subprocess_exec(
                "tasklist", "/fo", "csv", "/nh",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )
            stdout, _ = await asyncio.wait_for(proc.communicate(), timeout=5)
            apps = set()
            for line in stdout.decode(errors="replace").splitlines():
                parts = line.strip('"').split('","')
                if parts:
                    apps.add(parts[0].replace(".exe", ""))
            return sorted(apps)[:30]
        except Exception:
            return []


async def take_screenshot(display_only: bool = True) -> str | None:
    """Take a screenshot and return base64-encoded PNG. Tries multiple methods."""
    # Method 1: mss (fastest, no dependencies beyond mss)
    try:
        import mss
        import mss.tools

        with mss.mss() as sct:
            monitor = sct.monitors[1] if display_only and len(sct.monitors) > 1 else sct.monitors[0]
            img = sct.grab(monitor)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                tmp_path = f.name
            mss.tools.to_png(img.rgb, img.size, output=tmp_path)
            data = Path(tmp_path).read_bytes()
            Path(tmp_path).unlink(missing_ok=True)
            log.info(f"Screenshot captured via mss: {len(data)} bytes")
            return base64.b64encode(data).decode()
    except ImportError:
        pass
    except Exception as e:
        log.warning(f"mss screenshot failed: {e}")

    # Method 2: Pillow ImageGrab
    try:
        from PIL import ImageGrab
        img = ImageGrab.grab()
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            tmp_path = f.name
        img.save(tmp_path, "PNG")
        data = Path(tmp_path).read_bytes()
        Path(tmp_path).unlink(missing_ok=True)
        log.info(f"Screenshot captured via Pillow: {len(data)} bytes")
        return base64.b64encode(data).decode()
    except ImportError:
        pass
    except Exception as e:
        log.warning(f"Pillow screenshot failed: {e}")

    # Method 3: PowerShell fallback
    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            tmp_path = f.name.replace("\\", "\\\\")
        ps = (
            "Add-Type -AssemblyName System.Windows.Forms,System.Drawing; "
            "$s = [System.Windows.Forms.Screen]::PrimaryScreen; "
            "$b = New-Object System.Drawing.Bitmap($s.Bounds.Width,$s.Bounds.Height); "
            "$g = [System.Drawing.Graphics]::FromImage($b); "
            "$g.CopyFromScreen($s.Bounds.Location,[System.Drawing.Point]::Empty,$s.Bounds.Size); "
            f"$b.Save('{tmp_path}'); $g.Dispose(); $b.Dispose()"
        )
        proc = await asyncio.create_subprocess_exec(
            "powershell", "-NoProfile", "-Command", ps,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.DEVNULL,
        )
        await asyncio.wait_for(proc.communicate(), timeout=10)
        p = Path(tmp_path.replace("\\\\", "\\"))
        if p.exists():
            data = p.read_bytes()
            p.unlink(missing_ok=True)
            return base64.b64encode(data).decode()
    except Exception as e:
        log.warning(f"PowerShell screenshot failed: {e}")

    return None


async def describe_screen(anthropic_client) -> str:
    """Describe what's on the user's screen using vision or window list."""
    screenshot_b64 = await take_screenshot()
    if screenshot_b64 and anthropic_client:
        try:
            response = await anthropic_client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=300,
                system=(
                    "You are JARVIS analyzing a screenshot of the user's Windows desktop. "
                    "Describe what you see concisely: which apps are open, what the user "
                    "appears to be working on, any notable content visible. "
                    "Be specific about app names, file names, URLs, code, or documents visible. "
                    "2-4 sentences max. No markdown."
                ),
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": screenshot_b64,
                            },
                        },
                        {"type": "text", "text": "What's on my screen right now?"},
                    ],
                }],
            )
            return response.content[0].text
        except Exception as e:
            log.warning(f"Vision call failed, falling back to window list: {e}")

    windows = await get_active_windows()
    apps = await get_running_apps()

    if not windows and not apps:
        return "I wasn't able to see your screen, sir. You may need to install mss or Pillow."

    context_parts = []
    if windows:
        for w in windows[:15]:
            marker = " (ACTIVE)" if w["frontmost"] else ""
            context_parts.append(f"{w['app']}: {w['title']}{marker}")
    if apps:
        window_apps = {w["app"] for w in windows} if windows else set()
        bg = [a for a in apps if a not in window_apps]
        if bg:
            context_parts.append(f"Background: {', '.join(bg[:8])}")

    if anthropic_client and context_parts:
        try:
            response = await anthropic_client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=100,
                system=(
                    "You are JARVIS. Given the user's open windows and apps on Windows 11, "
                    "summarize what they appear to be working on in 1-2 sentences. "
                    "Natural voice, no markdown."
                ),
                messages=[{"role": "user", "content": "Open windows:\n" + "\n".join(context_parts)}],
            )
            return response.content[0].text
        except Exception:
            pass

    if windows:
        active = next((w for w in windows if w["frontmost"]), None)
        result = f"You have {len(windows)} windows open across {len({w['app'] for w in windows})} apps."
        if active:
            result += f" Currently focused on {active['app']}: {active['title']}."
        return result

    return f"Running apps: {', '.join(list(apps)[:10])}."


def format_windows_for_context(windows: list[dict]) -> str:
    """Format window list as context string for the LLM."""
    if not windows:
        return ""
    lines = ["Currently open on your desktop:"]
    for w in windows[:15]:
        marker = " (active)" if w["frontmost"] else ""
        lines.append(f"  - {w['app']}: {w['title']}{marker}")
    return "\n".join(lines)
