import time
import json
import logging

import win32gui
import win32process
import psutil

import uiautomation as uia

import mss
import numpy as np
import cv2
import pytesseract

from config import TESSERACT_CMD, OCR_CONFIDENCE_THRESHOLD, MAX_UIA_CHILDREN

log = logging.getLogger("vavi.probe")

pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


NOISY_CONTAINER_TYPES = {
    "MenuControl",
    "PaneControl",
    "GroupControl",
    "ToolBarControl",
    "CustomControl",
}

INTERESTING_TYPES = {
    "ButtonControl",
    "EditControl",
    "ComboBoxControl",
    "ListControl",
    "ListItemControl",
    "MenuItemControl",
    "TabItemControl",
    "CheckBoxControl",
    "RadioButtonControl",
    "HyperlinkControl",
    "DocumentControl",
    "TreeControl",
    "TreeItemControl",
}


def _safe_str(x):
    return (x or "").strip()


def is_useful_control(c, focused=None):
    try:
        ctype = _safe_str(c.ControlTypeName)
        name = _safe_str(c.Name)
        cls = _safe_str(c.ClassName)

        if focused is not None and c == focused:
            return True

        if name:
            return True

        if ctype in INTERESTING_TYPES:
            return True

        if ctype in NOISY_CONTAINER_TYPES and not name and not cls:
            return False

        return False

    except Exception:
        return False


def control_key(c):
    try:
        return (_safe_str(c.ControlTypeName), _safe_str(c.Name), _safe_str(c.ClassName))
    except Exception:
        return ("", "", "")


def get_foreground_window_info(hwnd=None):
    if hwnd is None:
        hwnd = win32gui.GetForegroundWindow()
    title = win32gui.GetWindowText(hwnd)
    rect = win32gui.GetWindowRect(hwnd)

    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    proc = psutil.Process(pid)

    return {
        "hwnd": int(hwnd),
        "title": title,
        "bounds": {"left": rect[0], "top": rect[1], "right": rect[2], "bottom": rect[3]},
        "process": {"pid": pid, "name": proc.name(), "exe": proc.exe()},
    }


def get_uia_summary(max_children=MAX_UIA_CHILDREN, hwnd=None):
    if hwnd is not None:
        fg = uia.ControlFromHandle(hwnd)
    else:
        fg = uia.GetForegroundControl()
    focused = uia.GetFocusedControl()

    children = []
    seen = set()

    try:
        for c in fg.GetChildren():
            if not is_useful_control(c, focused=focused):
                continue

            key = control_key(c)
            if key in seen:
                continue
            seen.add(key)

            try:
                auto_id = _safe_str(c.AutomationId)
            except Exception:
                auto_id = ""

            children.append({
                "type": c.ControlTypeName,
                "name": c.Name,
                "class": c.ClassName,
                "automation_id": auto_id,
            })

            if len(children) >= max_children:
                break

    except Exception as e:
        log.warning("Error iterating UIA children: %s", e)

    return {
        "foreground": {"name": fg.Name, "class": fg.ClassName, "type": fg.ControlTypeName},
        "focused": {"name": focused.Name, "class": focused.ClassName, "type": focused.ControlTypeName},
        "children_preview": children,
    }


def _valid_bounds(bounds):
    w = bounds.get("right", 0) - bounds.get("left", 0)
    h = bounds.get("bottom", 0) - bounds.get("top", 0)
    return w > 0 and h > 0


def _capture_and_threshold(bounds):
    if not _valid_bounds(bounds):
        return None

    left = bounds["left"]
    top = bounds["top"]
    width = max(1, bounds["right"] - bounds["left"])
    height = max(1, bounds["bottom"] - bounds["top"])

    with mss.mss() as sct:
        shot = sct.grab({"left": left, "top": top, "width": width, "height": height})
        img = np.array(shot)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return gray


def ocr_window_region(bounds, _gray=None):
    if _gray is None:
        _gray = _capture_and_threshold(bounds)
    if _gray is None:
        log.warning("OCR skipped: zero-size window bounds")
        return ""

    text = pytesseract.image_to_string(_gray)
    text = "\n".join([line.strip() for line in text.splitlines() if line.strip()])
    return text


def ocr_window_items(bounds, _gray=None):
    if _gray is None:
        _gray = _capture_and_threshold(bounds)
    if _gray is None:
        log.warning("OCR items skipped: zero-size window bounds")
        return []

    data = pytesseract.image_to_data(_gray, output_type=pytesseract.Output.DICT)

    items = []
    n = len(data["text"])
    for i in range(n):
        txt = (data["text"][i] or "").strip()
        conf = data["conf"][i]
        if not txt:
            continue
        try:
            conf_i = int(conf)
        except Exception:
            conf_i = -1

        if conf_i < OCR_CONFIDENCE_THRESHOLD:
            continue

        items.append({
            "text": txt,
            "conf": conf_i,
            "x": int(data["left"][i]),
            "y": int(data["top"][i]),
            "w": int(data["width"][i]),
            "h": int(data["height"][i]),
        })

    return items


def capture_desktop_state(do_ocr=False, target_hwnd=None):
    state = {}
    state["window"] = get_foreground_window_info(hwnd=target_hwnd)

    try:
        state["uia"] = get_uia_summary(hwnd=target_hwnd)
    except Exception as e:
        log.error("UIA failed: %s", e)
        state["uia"] = {}
        state["uia_error"] = str(e)

    if do_ocr:
        try:
            bounds = state["window"]["bounds"]
            gray = _capture_and_threshold(bounds)
            state["ocr_text"] = ocr_window_region(bounds, _gray=gray)
            state["ocr_items"] = ocr_window_items(bounds, _gray=gray)
        except Exception as e:
            log.error("OCR failed: %s", e)
            state["ocr_error"] = str(e)

    state["timestamp"] = time.time()
    return state


if __name__ == "__main__":
    import config 
    print("Desktop probe running. Ctrl+C to stop.")
    while True:
        state = capture_desktop_state(do_ocr=True)
        print(json.dumps(state, indent=2))
        time.sleep(5.0)
        print("-" * 80)
