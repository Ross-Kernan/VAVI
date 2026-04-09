import logging

import pyautogui
import win32gui
import win32con

log = logging.getLogger("vavi.commands.input_actions")

_WINDOW_COMMANDS = {
    "minimize": "minimize",
    "minimise": "minimize",
    "maximize": "maximize",
    "maximise": "maximize",
    "close":    "close",
    "restore":  "restore",
}


def try_window_command(target: str, hwnd: int = None) -> bool:
    action = _WINDOW_COMMANDS.get(target.strip().lower())
    if action is None:
        return False

    if hwnd is None:
        hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return False

    if action == "minimize":
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    elif action == "maximize":
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    elif action == "restore":
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    elif action == "close":
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    return True


def press_key(key: str) -> bool:
    raw = key.strip().lower()
    if not raw:
        return False

    _MULTI = [
        ("page up",    "pageup"),
        ("page down",  "pagedown"),
        ("back space", "backspace"),
        ("f one",      "f1"),  ("f two",    "f2"),  ("f three", "f3"),
        ("f four",     "f4"),  ("f five",   "f5"),  ("f six",   "f6"),
        ("f seven",    "f7"),  ("f eight",  "f8"),  ("f nine",  "f9"),
        ("f ten",      "f10"), ("f eleven", "f11"), ("f twelve","f12"),
    ]
    for phrase, replacement in _MULTI:
        raw = raw.replace(phrase, replacement)

    _WORD = {
        "control":  "ctrl",
        "windows":  "win",
        "window":   "win",
        "super":    "win",
        "return":   "enter",
        "escape":   "esc",
        "spacebar": "space",
        "del":      "delete",
        "insert":   "insert",
    }
    tokens = [_WORD.get(w, w) for w in raw.split()]

    _VALID = {
        "ctrl", "alt", "shift", "win",
        "enter", "tab", "esc", "backspace", "delete", "insert",
        "home", "end", "pageup", "pagedown",
        "up", "down", "left", "right",
        "space", "printscreen", "scrolllock", "pause",
        *[f"f{i}" for i in range(1, 13)],
        *"abcdefghijklmnopqrstuvwxyz",
        *[str(i) for i in range(10)],
    }
    if not tokens or not all(t in _VALID for t in tokens):
        log.debug("press_key: unrecognised key(s) in %r → tokens %s", key, tokens)
        return False

    try:
        if len(tokens) == 1:
            pyautogui.press(tokens[0])
        else:
            pyautogui.hotkey(*tokens)
        return True
    except Exception as e:
        log.debug("press_key failed: %s", e)
        return False
