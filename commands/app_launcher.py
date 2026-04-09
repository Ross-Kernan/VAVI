import subprocess
import shutil
import logging

import pyautogui
import win32gui
import uiautomation as uia

from commands.uia_utils import iter_descendants

log = logging.getLogger("vavi.commands.app_launcher")

_APP_ALIASES = {
    "chrome":               "chrome",
    "google chrome":        "chrome",
    "google":               "chrome",
    "firefox":              "firefox",
    "firefox browser":      "firefox",
    "edge":                 "msedge",
    "microsoft edge":       "msedge",
    "notepad":              "notepad",
    "explorer":             "explorer",
    "file explorer":        "explorer",
    "this pc":              "explorer",
    "file manager":         "explorer",
    "calculator":           "calc",
    "paint":                "mspaint",
    "word":                 "winword",
    "excel":                "excel",
    "powerpoint":           "powerpnt",
    "visual studio code":   "code",
    "vs code":              "code",
    "task manager":         "taskmgr",
    "taskmgr":              "taskmgr",
    "discord":              "discord",
    "spotify":              "spotify",
    "zoom":                 "zoom",
    "teams":                "ms-teams",
    "microsoft teams":      "ms-teams",
    "obs":                  "obs64",
    "obs studio":           "obs64",
    "vlc":                  "vlc",
    "slack":                "slack",
    "snipping tool":        "SnippingTool",
    "snip":                 "SnippingTool",
    "cmd":                  "cmd",
    "command prompt":       "cmd",
    "powershell":           "powershell",
    "terminal":             "wt",
    "windows terminal":     "wt",
    "outlook":              "outlook",
    "skype":                "skype",
}

# UWP / Store apps that can't be launched via a plain exe name
_UWP_PROTOCOLS = {
    "photos":               "ms-photos:",
    "microsoft photos":     "ms-photos:",
    "settings":             "ms-settings:",
    "windows settings":     "ms-settings:",
    "store":                "ms-windows-store:",
    "microsoft store":      "ms-windows-store:",
    "calendar":             "outlookcal:",
    "mail":                 "outlookmail:",
    "camera":               "microsoft.windows.camera:",
    "weather":              "msnweather:",
    "clock":                "ms-clock:",
    "alarms":               "ms-clock:",
    "sticky notes":         "ms-stickynotes:",
    "xbox":                 "xbox:",
    "game bar":             "ms-gamebar:",
    "movies":               "mswindowsvideo:",
    "movies and tv":        "mswindowsvideo:",
}


def _click_taskbar_app(name: str) -> bool:
    # Search the Windows taskbar UIA tree for a button matching *name* and click it
    hwnd = win32gui.FindWindow("Shell_TrayWnd", None)
    if not hwnd:
        log.warning("Taskbar window not found")
        return False

    name_lower = name.strip().lower()
    try:
        taskbar = uia.ControlFromHandle(hwnd)
        for c in iter_descendants(taskbar, max_nodes=300):
            try:
                c_name = (c.Name or "").strip().lower()
                ctype = (c.ControlTypeName or "").strip()
                if not c_name:
                    continue
                if name_lower in c_name:
                    if ctype in ("ButtonControl", "ListItemControl"):
                        try:
                            c.Click()
                            return True
                        except Exception:
                            r = c.BoundingRectangle
                            cx = int((r.left + r.right) / 2)
                            cy = int((r.top + r.bottom) / 2)
                            if cx > 0 or cy > 0:
                                pyautogui.click(cx, cy)
                                return True
            except Exception:
                continue
    except Exception as e:
        log.warning("Taskbar search failed: %s", e)

    return False


_WEBSITE_ALIASES = {
    "youtube":      "https://www.youtube.com",
    "gmail":        "https://mail.google.com",
    "facebook":     "https://www.facebook.com",
    "reddit":       "https://www.reddit.com",
    "github":       "https://www.github.com",
    "netflix":      "https://www.netflix.com",
    "amazon":       "https://www.amazon.com",
    "wikipedia":    "https://www.wikipedia.org",
}


def launch_application(name: str) -> bool:
    # Launch an application by name, alias, or website. Returns True if launched
    name_lower = name.strip().lower()

    url = _WEBSITE_ALIASES.get(name_lower)
    if url:
        try:
            subprocess.Popen(f'start "" "{url}"', shell=True)
            log.info("Opened website '%s' -> %s", name_lower, url)
            return True
        except Exception as e:
            log.warning("Failed to open URL %s: %s", url, e)

    protocol = _UWP_PROTOCOLS.get(name_lower)
    if protocol:
        try:
            subprocess.Popen(f'start "" "{protocol}"', shell=True)
            log.info("Opened UWP app '%s' -> %s", name_lower, protocol)
            return True
        except Exception as e:
            log.warning("Failed to open UWP protocol %s: %s", protocol, e)

    exe = _APP_ALIASES.get(name_lower)

    if exe is None:
        if _click_taskbar_app(name_lower):
            log.info("Launched '%s' via taskbar click", name_lower)
            return True
        exe = name_lower.split()[0]

    path = shutil.which(exe)
    if path:
        try:
            subprocess.Popen([path])
            log.info("Launched '%s' via PATH: %s", exe, path)
            return True
        except Exception as e:
            log.warning("Popen failed for %s: %s", path, e)

    try:
        subprocess.Popen(f'start "" "{exe}"', shell=True)
        log.info("Launched '%s' via shell start", exe)
        return True
    except Exception as e:
        log.warning("Shell start failed for %s: %s", exe, e)

    return _click_taskbar_app(name_lower)
