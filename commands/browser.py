import re
import time
import logging
from typing import Tuple

import pyautogui
import win32gui
import uiautomation as uia

from config import MAX_SCAN_DEPTH
from commands.uia_utils import iter_descendants, find_control_by_name, try_invoke_or_click

log = logging.getLogger("vavi.commands.browser")


def _norm(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s


def click_text_via_ocr(state: dict, target: str, double: bool = False, right: bool = False) -> bool:
    items = state.get("ocr_items") or []
    if not items:
        return False

    bounds = (state.get("window") or {}).get("bounds") or {}
    win_left = int(bounds.get("left", 0))
    win_top = int(bounds.get("top", 0))

    tgt = _norm(target)
    if not tgt:
        return False

    def _do_click(x, y, w, h):
        cx = win_left + x + w // 2
        cy = win_top  + y + h // 2
        if double:
            pyautogui.doubleClick(cx, cy)
        elif right:
            pyautogui.rightClick(cx, cy)
        else:
            pyautogui.click(cx, cy)

    for it in items:
        if it.get("conf", -1) < 35:
            continue
        if _norm(it["text"]) == tgt:
            _do_click(it["x"], it["y"], it["w"], it["h"])
            return True

    tgt_compact = tgt.replace(" ", "")
    for it in items:
        if it.get("conf", -1) < 35:
            continue
        if _norm(it["text"]).replace(" ", "") == tgt_compact:
            _do_click(it["x"], it["y"], it["w"], it["h"])
            return True

    tgt_words = tgt.split()
    n = len(tgt_words)
    if n >= 2:
        valid = [it for it in items if it.get("conf", -1) >= 35 and _norm(it["text"])]
        valid.sort(key=lambda it: (it["y"], it["x"]))

        for i in range(len(valid) - n + 1):
            window = valid[i : i + n]
            ys = [w["y"] for w in window]
            if max(ys) - min(ys) > 15:
                continue
            if [_norm(w["text"]) for w in window] == tgt_words:
                x1 = min(w["x"] for w in window)
                y1 = min(w["y"] for w in window)
                x2 = max(w["x"] + w["w"] for w in window)
                y2 = max(w["y"] + w["h"] for w in window)
                _do_click(x1, y1, x2 - x1, y2 - y1)
                return True

    return False


def click_browser_result(n: int, target_hwnd=None) -> Tuple[bool, str]:
    if target_hwnd is not None:
        fg = uia.ControlFromHandle(target_hwnd)
        try:
            win_rect = win32gui.GetWindowRect(target_hwnd)
        except Exception:
            win_rect = None
    else:
        fg = uia.GetForegroundControl()
        win_rect = None

    chrome_bottom = 0
    if win_rect:
        chrome_bottom = win_rect[1] + int((win_rect[3] - win_rect[1]) * 0.12)

    candidates = []
    for c in iter_descendants(fg, max_nodes=MAX_SCAN_DEPTH):
        try:
            ctype = (c.ControlTypeName or "").strip()
            if ctype != "HyperlinkControl":
                continue
            c_name = (c.Name or "").strip()
            if not c_name or len(c_name) < 5:
                continue

            c_name_lower = c_name.lower()
            if "https://" not in c_name_lower and "http://" not in c_name_lower:
                continue

            try:
                r = c.BoundingRectangle
                if r.right <= r.left or r.bottom <= r.top:
                    continue
                link_w = r.right - r.left
                link_h = r.bottom - r.top
                if link_w < 80 or link_h < 10:
                    continue
                cx = (r.left + r.right) // 2
                cy = (r.top + r.bottom) // 2
            except Exception:
                continue

            if win_rect:
                if not (win_rect[0] <= cx <= win_rect[2] and
                        win_rect[1] <= cy <= win_rect[3]):
                    continue

            if chrome_bottom and r.top < chrome_bottom:
                continue

            candidates.append((r.top, r.left, c, c_name))
        except Exception:
            continue

    if not candidates:
        return False, "I couldn't find any results to click."

    candidates.sort(key=lambda x: (x[0], x[1]))

    deduped: list = []
    last_top = None
    for item in candidates:
        if last_top is None or abs(item[0] - last_top) > 5:
            deduped.append(item)
            last_top = item[0]

    total = len(deduped)
    if n < 1 or n > total:
        return False, f"I only found {total} result{'s' if total != 1 else ''}."

    _, _, control, name = deduped[n - 1]
    try:
        r = control.BoundingRectangle
        x = int((r.left + r.right) / 2)
        y = int((r.top + r.bottom) / 2)
        uia.Click(x, y)

        display = re.sub(r'\s*https?://\S+', '', name).strip()
        if not display:
            display = f"result {n}"
        elif len(display) > 60:
            display = display[:57] + "..."
        log.info("Clicked browser result %d: %s", n, display)
        return True, f"Clicked result {n}: {display}."
    except Exception as e:
        log.error("click_browser_result failed: %s", e)
        return False, f"I found result {n} but couldn't click it."


def _find_page_search_control(target_hwnd=None):
    root = uia.ControlFromHandle(target_hwnd) if target_hwnd else uia.GetForegroundControl()

    win_rect = None
    if target_hwnd:
        try:
            win_rect = win32gui.GetWindowRect(target_hwnd)
        except Exception:
            pass

    doc = None
    for c in iter_descendants(root, max_nodes=120):
        try:
            if (c.ControlTypeName or "").strip() == "DocumentControl":
                doc = c
                break
        except Exception:
            continue

    scan_root = doc if doc else root

    search_ctrl = None
    first_edit  = None

    for c in iter_descendants(scan_root, max_nodes=MAX_SCAN_DEPTH):
        try:
            ctype = (c.ControlTypeName or "").strip()
            if ctype not in ("EditControl", "ComboBoxControl"):
                continue

            c_name = (c.Name or "").strip()
            c_name_lower = c_name.lower()

            if "address" in c_name_lower:
                continue

            try:
                r = c.BoundingRectangle
                if (r.right - r.left) < 50 or (r.bottom - r.top) < 5:
                    continue
                if win_rect:
                    cx = (r.left + r.right) / 2
                    cy = (r.top + r.bottom) / 2
                    if not (win_rect[0] <= cx <= win_rect[2] and
                            win_rect[1] <= cy <= win_rect[3]):
                        log.debug(
                            "_find_page_search_control: skipping '%s' — centre "
                            "(%.0f,%.0f) outside window %s",
                            c_name, cx, cy, win_rect,
                        )
                        continue
            except Exception:
                continue

            if "search" in c_name_lower:
                search_ctrl = c
                break

            if first_edit is None:
                first_edit = c
        except Exception:
            continue

    return search_ctrl or first_edit


def perform_search(query: str, state: dict, target_hwnd=None):
    window_title = ((state.get("window") or {}).get("title") or "").lower()

    on_youtube = "youtube" in window_title
    in_browser = any(b in window_title for b in ("chrome", "firefox", "edge", "mozilla"))

    if on_youtube:
        try:
            pyautogui.press("escape")
            time.sleep(0.15)
            pyautogui.press("/")
            time.sleep(0.3)
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.1)
            uia.SendKeys(query, waitTime=0)
            time.sleep(0.1)
            uia.SendKeys("{Enter}", waitTime=0)
            log.info("YouTube search: %s", query)
            return True, f"Searching YouTube for {query}."
        except Exception as e:
            log.warning("YouTube search shortcut failed: %s", e)

    on_new_tab = "new tab" in window_title or window_title in ("mozilla firefox", "google chrome", "microsoft edge")

    if in_browser and not on_youtube and not on_new_tab:
        page_ctrl = _find_page_search_control(target_hwnd)

        if page_ctrl is not None:
            try:
                r = page_ctrl.BoundingRectangle
                px = int((r.left + r.right) / 2)
                py = int((r.top + r.bottom) / 2)
                uia.Click(px, py)
                time.sleep(0.25)
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.1)
                uia.SendKeys(query, waitTime=0)
                time.sleep(0.1)
                uia.SendKeys("{Enter}", waitTime=0)
                log.info("Page search control (%s): %s",
                         (page_ctrl.Name or ""), query)
                return True, f"Searching for {query}."
            except Exception as e:
                log.warning("Page search control click failed: %s", e)

    if in_browser or on_youtube:
        try:
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.3)
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.1)
            uia.SendKeys(query, waitTime=0)
            time.sleep(0.1)
            uia.SendKeys("{Enter}", waitTime=0)
            log.info("Browser address bar search: %s", query)
            return True, f"Searching for {query}."
        except Exception as e:
            log.warning("Browser address bar search failed: %s", e)

    win_rect = None
    if target_hwnd:
        try:
            win_rect = win32gui.GetWindowRect(target_hwnd)
        except Exception:
            pass

    control = find_control_by_name("search bar", max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)
    if control is None:
        control = find_control_by_name("search", max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)

    if control:
        if try_invoke_or_click(control, win_rect=win_rect):
            time.sleep(0.2)
            uia.SendKeys(query, waitTime=0)
            time.sleep(0.1)
            uia.SendKeys("{Enter}", waitTime=0)
            log.info("UIA search: %s", query)
            return True, f"Searching for {query}."

    return False, f"I couldn't find a search bar to search for {query}."


# ---------------------------------------------------------------------------
# Tab management
# ---------------------------------------------------------------------------

_TAB_HOTKEYS = {
    "new":      ("ctrl", "t"),
    "close":    ("ctrl", "w"),
    "next":     ("ctrl", "tab"),
    "right":    ("ctrl", "tab"),
    "previous": ("ctrl", "shift", "tab"),
    "prev":     ("ctrl", "shift", "tab"),
    "back":     ("ctrl", "shift", "tab"),
    "left":     ("ctrl", "shift", "tab"),
    "last":     ("ctrl", "shift", "tab"),
}

_TAB_MESSAGES = {
    "new":      "New tab opened.",
    "close":    "Tab closed.",
    "next":     "Next tab.",
    "right":    "Next tab.",
    "previous": "Previous tab.",
    "prev":     "Previous tab.",
    "back":     "Previous tab.",
    "left":     "Previous tab.",
    "last":     "Previous tab.",
}

_SPOKEN_TAB_NUMBERS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9,
}


def handle_tab_command(arg: str):
    arg = (arg or "").strip().lower()

    if arg in _TAB_HOTKEYS:
        try:
            pyautogui.hotkey(*_TAB_HOTKEYS[arg])
            return True, _TAB_MESSAGES[arg]
        except Exception as e:
            log.warning("Tab hotkey failed for '%s': %s", arg, e)
            return False, f"I couldn't perform the tab action."

    m = re.match(r"^close\s+(.+)$", arg)
    if m:
        num_str = m.group(1).strip()
        n = None
        try:
            n = int(num_str)
        except ValueError:
            n = _SPOKEN_TAB_NUMBERS.get(num_str)
        if n is not None:
            try:
                key = str(min(n, 9))
                pyautogui.hotkey("ctrl", key)
                time.sleep(0.25)
                pyautogui.hotkey("ctrl", "w")
                return True, f"Tab {n} closed."
            except Exception as e:
                log.warning("Close-tab-number failed: %s", e)
                return False, "I couldn't close that tab."

    n = None
    try:
        n = int(arg)
    except ValueError:
        n = _SPOKEN_TAB_NUMBERS.get(arg)

    if n is not None:
        if 1 <= n <= 8:
            try:
                pyautogui.hotkey("ctrl", str(n))
                return True, f"Switched to tab {n}."
            except Exception as e:
                log.warning("Tab number hotkey failed: %s", e)
        elif n == 9:
            try:
                pyautogui.hotkey("ctrl", "9")
                return True, "Switched to last tab."
            except Exception as e:
                log.warning("Tab 9 hotkey failed: %s", e)

    return False, f"I didn't understand tab command: {arg}."
