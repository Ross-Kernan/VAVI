import logging
import time

import pyautogui
import win32gui

log = logging.getLogger("vavi.commands.scroll")


def _focus_target_window(hwnd: int) -> bool:
    if not hwnd or not win32gui.IsWindow(hwnd):
        return False
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.15)
        return True
    except Exception as e:
        log.debug("SetForegroundWindow(%s) failed: %s", hwnd, e)
        return False


def scroll_page(direction: str, clicks: int = 5, hwnd=None) -> bool:
    try:
        if hwnd is None:
            hwnd = win32gui.GetForegroundWindow()
        _focus_target_window(hwnd)
        key = "pagedown" if direction == "down" else "pageup"
        presses = max(1, clicks // 5)
        for _ in range(presses):
            pyautogui.press(key)
        log.info("Scrolled %s %d press(es) via %s", direction, presses, key)
        return True
    except Exception as e:
        log.error("scroll_page failed: %s", e)
        return False


def scroll_lines(direction: str, lines: int = 3, hwnd=None) -> bool:
    try:
        if hwnd is None:
            hwnd = win32gui.GetForegroundWindow()
        _focus_target_window(hwnd)
        try:
            rect = win32gui.GetWindowRect(hwnd)
            cx = (rect[0] + rect[2]) // 2
            cy = (rect[1] + rect[3]) // 2
            pyautogui.moveTo(cx, cy)
        except Exception:
            pass
        amount = lines if direction == "up" else -lines
        pyautogui.scroll(amount)
        log.info("Scrolled %s %d line(s) via wheel", direction, lines)
        return True
    except Exception as e:
        log.error("scroll_lines failed: %s", e)
        return False
