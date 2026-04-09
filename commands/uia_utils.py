import re
import logging
from difflib import get_close_matches

import pyautogui
import win32gui
import win32con
import win32process
import uiautomation as uia

from config import MAX_SCAN_DEPTH

log = logging.getLogger("vavi.commands.uia_utils")

_BROWSER_SUFFIXES = (
    " \u2014 Mozilla Firefox",
    " - Mozilla Firefox",
    " \u2014 Google Chrome",
    " - Google Chrome",
    " \u2014 Microsoft Edge",
    " - Microsoft Edge",
    " \u2014 Chromium",
    " - Chromium",
)

_NOISE_PHRASES = [
    "the editor is not accessible at this time",
    "to enable screen reader optimized mode",
    "screen reader optimized mode",
    "shift+alt+f1",
]


def describe_focus(state: dict) -> str:
    def clean_text(s: str) -> str:
        s = (s or "").strip()
        s_low = s.lower()
        for p in _NOISE_PHRASES:
            if p in s_low:
                return ""
        return s

    def strip_browser(s: str) -> str:
        for suffix in _BROWSER_SUFFIXES:
            if s.endswith(suffix):
                return s[: -len(suffix)].strip()
        return s

    u = state.get("uia") or {}
    fg = u.get("foreground") or {}
    focused = u.get("focused") or {}

    fg_name = strip_browser(clean_text(fg.get("name") or ""))
    window_title = strip_browser(clean_text((state.get("window") or {}).get("title") or ""))
    focused_name = clean_text(focused.get("name") or "")

    app_part = fg_name or window_title or "Active window"

    if not focused_name or focused_name.lower() == app_part.lower():
        return f"{app_part}."

    return f"{app_part}. Focus on {focused_name}."


def normalize_window_title(title: str) -> str:
    t = (title or "").strip()
    if t.startswith("*"):
        t = t[1:].strip()
    return t


def summarize_actionables(state: dict, limit: int = 8) -> str:
    u = state.get("uia") or {}
    children = u.get("children_preview") or []

    keep_types = {
        "ButtonControl", "HyperlinkControl", "MenuItemControl", "TabItemControl",
        "CheckBoxControl", "RadioButtonControl", "ComboBoxControl", "ListItemControl"
    }

    names = []
    for c in children:
        ctype = (c.get("type") or "").strip()
        name = (c.get("name") or "").strip()
        if not name:
            name = (c.get("automation_id") or "").strip()
        if not name:
            continue
        if ctype in keep_types:
            names.append(name)

    seen = set()
    names = [n for n in names if not (n.lower() in seen or seen.add(n.lower()))]

    if not names:
        return "I can't see any named clickable items in the preview."
    names = names[:limit]
    return "Clickable items include: " + ", ".join(names) + "."


# ---------------------------------------------------------------------------
# UIA tree traversal
# ---------------------------------------------------------------------------
def iter_descendants(root, max_nodes: int = 300):
    q = [root]
    count = 0
    while q and count < max_nodes:
        cur = q.pop(0)
        yield cur
        count += 1
        try:
            q.extend(cur.GetChildren())
        except Exception:
            pass


_PREFERRED_CLICK_TYPES = {
    "HyperlinkControl", "ButtonControl", "MenuItemControl",
    "ListItemControl", "LinkControl",
}

_ALL_INTERACTIVE_TYPES = {
    "ButtonControl", "HyperlinkControl", "MenuItemControl", "TabItemControl",
    "ListItemControl", "CheckBoxControl", "RadioButtonControl", "ComboBoxControl",
    "EditControl", "TreeItemControl", "SplitButtonControl", "LinkControl",
    "ToggleControl", "DataItemControl",
}


_INPUT_SUFFIX_RE = re.compile(r'\s+(bar|box|field|input)$')
_INPUT_PREFERRED_TYPES = {"EditControl", "ComboBoxControl"}

_CLICK_TARGET_ALIASES = {
    "messages": "message",
    "msg": "message",
    "msgs": "message",
}


def find_control_by_name(name: str, max_nodes: int = MAX_SCAN_DEPTH, hwnd=None):
    target = name.strip().lower()

    target = _CLICK_TARGET_ALIASES.get(target, target)

    target = re.sub(r'\s+dot\s+', '.', target)

    prefer_input = bool(_INPUT_SUFFIX_RE.search(target))
    original_target = target
    if prefer_input:
        target = _INPUT_SUFFIX_RE.sub("", target).strip()

    if hwnd is not None:
        fg = uia.ControlFromHandle(hwnd)
        try:
            win_rect = win32gui.GetWindowRect(hwnd)
        except Exception:
            win_rect = None
    else:
        fg = uia.GetForegroundControl()
        win_rect = None

    interactive_candidates = []
    fallback_candidates = []
    _fuzzy_pool: list = []

    for c in iter_descendants(fg, max_nodes=max_nodes):
        try:
            c_name = (c.Name or "").strip()
            try:
                c_auto_id = (c.AutomationId or "").strip()
            except Exception:
                c_auto_id = ""
            try:
                c_class = (c.ClassName or "").strip()
            except Exception:
                c_class = ""

            labels = [l for l in (c_name, c_auto_id, c_class) if l]
            if not labels:
                continue

            try:
                ctype = (c.ControlTypeName or "").strip()
            except Exception:
                ctype = ""

            if prefer_input:
                is_preferred = ctype in _INPUT_PREFERRED_TYPES
            else:
                is_preferred = ctype in _PREFERRED_CLICK_TYPES

            is_interactive = ctype in _ALL_INTERACTIVE_TYPES

            if is_interactive and c_name:
                _fuzzy_pool.append((c_name.lower(), c))

            _phrase_matched = False
            if prefer_input and original_target and ctype in _INPUT_PREFERRED_TYPES:
                for label in labels:
                    label_l = label.lower()
                    if original_target not in label_l:
                        continue
                    orig_score = len(original_target) / len(label_l) if label_l else 0
                    if orig_score < 0.2:
                        continue
                    if win_rect:
                        try:
                            r = c.BoundingRectangle
                            if r.right > r.left and r.bottom > r.top:
                                cx = (r.left + r.right) / 2
                                cy = (r.top + r.bottom) / 2
                                if not (win_rect[0] <= cx <= win_rect[2] and
                                        win_rect[1] <= cy <= win_rect[3]):
                                    log.debug(
                                        "Skipping phrase-match '%s' — centre outside window",
                                        label,
                                    )
                                    continue
                        except Exception:
                            pass
                    try:
                        _r = c.BoundingRectangle
                        _y_top = _r.top if _r.top > 0 else 9999
                        _x_left = _r.left if _r.left >= 0 else 9999
                    except Exception:
                        _y_top = 9999
                        _x_left = 9999
                    rank = (False, True, True, True, orig_score, -_y_top, -_x_left, 0)
                    interactive_candidates.append((rank, c))
                    _phrase_matched = True
                    break
            if _phrase_matched:
                continue

            for label in labels:
                label_l = label.lower()
                is_exact = (label_l == target)
                is_substring = (target in label_l)
                is_prefix = label_l.startswith(target)

                target_words = target.split()
                label_words_set = set(label_l.split())
                is_word_match = (
                    bool(target_words)
                    and not is_substring
                    and all(w in label_words_set for w in target_words)
                )

                if not is_exact and not is_substring and not is_word_match:
                    continue

                text_score = len(target) / len(label_l) if label_l else 0

                if not is_exact and not is_prefix and not is_word_match and text_score < 0.5:
                    continue

                if win_rect:
                    try:
                        r = c.BoundingRectangle
                        has_size = r.right > r.left and r.bottom > r.top
                        if has_size:
                            cx = (r.left + r.right) / 2
                            cy = (r.top + r.bottom) / 2
                            in_window = (win_rect[0] <= cx <= win_rect[2] and
                                         win_rect[1] <= cy <= win_rect[3])
                            if not in_window:
                                log.debug(
                                    "Skipping '%s' — centre (%.0f,%.0f) outside window",
                                    label, cx, cy,
                                )
                                continue
                    except Exception:
                        pass

                try:
                    _r = c.BoundingRectangle
                    _area = max(0, _r.right - _r.left) * max(0, _r.bottom - _r.top)
                    _y_top = _r.top if _r.top > 0 else 9999
                    _x_left = _r.left if _r.left >= 0 else 9999
                except Exception:
                    _area = 0
                    _y_top = 9999
                    _x_left = 9999
                if _area == 0:
                    log.debug("Skipping zero-bounds control '%s' (%s)", label, ctype)
                    continue

                is_nav_link = ctype == "HyperlinkControl" and (
                    "https://" in label_l or "http://" in label_l
                )

                # Priority: nav link -> match quality -> position (top/left first)
                rank = (is_nav_link, is_exact, is_prefix, is_preferred, text_score, -_y_top, -_x_left, _area)
                entry = (rank, c)
                if ctype in ("TabItemControl", "EditControl"):
                    fallback_candidates.append(entry)
                elif is_interactive:
                    interactive_candidates.append(entry)
                else:
                    fallback_candidates.append(entry)

        except Exception:
            continue

    if not interactive_candidates and not fallback_candidates and _fuzzy_pool:
        fuzzy_names = [n for n, _ in _fuzzy_pool]
        close = get_close_matches(target, fuzzy_names, n=1, cutoff=0.72)
        if close:
            matched_name = close[0]
            for fn, fc in _fuzzy_pool:
                if fn != matched_name:
                    continue
                try:
                    _r = fc.BoundingRectangle
                    _area = max(0, _r.right - _r.left) * max(0, _r.bottom - _r.top)
                    _y_top = _r.top if _r.top > 0 else 9999
                    _x_left = _r.left if _r.left >= 0 else 9999
                except Exception:
                    _area, _y_top, _x_left = 0, 9999, 9999
                if _area == 0:
                    continue
                try:
                    _fctype = (fc.ControlTypeName or "").strip()
                    _fpreferred = _fctype in _PREFERRED_CLICK_TYPES
                except Exception:
                    _fpreferred = False
                rank = (False, False, False, _fpreferred, 0.6, -_y_top, -_x_left, _area)
                interactive_candidates.append((rank, fc))
                log.debug("Fuzzy-matched '%s' → '%s'", target, matched_name)
                break

    pool = interactive_candidates if interactive_candidates else fallback_candidates
    if not pool:
        return None

    pool.sort(key=lambda x: x[0], reverse=True)
    return pool[0][1]


def try_invoke_or_click(control, win_rect=None) -> bool:
    if control is None:
        return False

    try:
        ctype = (control.ControlTypeName or "").strip()
    except Exception:
        ctype = ""

    def _coord_click() -> bool:
        try:
            r = control.BoundingRectangle
            x = int((r.left + r.right) / 2)
            y = int((r.top + r.bottom) / 2)
            if x == 0 and y == 0 and r.right == 0 and r.bottom == 0:
                return False
            if win_rect and not (win_rect[0] <= x <= win_rect[2] and
                                 win_rect[1] <= y <= win_rect[3]):
                log.warning(
                    "Coordinate click (%d,%d) is outside window bounds %s — skipping",
                    x, y, win_rect,
                )
                return False
            uia.Click(x, y)
            return True
        except Exception:
            return False

    if ctype == "HyperlinkControl":
        return _coord_click()

    try:
        invoke = getattr(control, "Invoke", None)
        if callable(invoke):
            invoke()
            return True
    except Exception:
        pass

    if ctype in _ALL_INTERACTIVE_TYPES:
        try:
            click_m = getattr(control, "Click", None)
            if callable(click_m):
                click_m()
                return True
        except Exception:
            pass

    return _coord_click()


def try_double_click(control, win_rect=None) -> bool:
    if control is None:
        return False
    try:
        r = control.BoundingRectangle
        x = int((r.left + r.right) / 2)
        y = int((r.top + r.bottom) / 2)
        if x == 0 and y == 0 and r.right == 0 and r.bottom == 0:
            return False
        if win_rect and not (win_rect[0] <= x <= win_rect[2] and
                             win_rect[1] <= y <= win_rect[3]):
            log.warning(
                "Double-click (%d,%d) is outside window bounds %s — skipping",
                x, y, win_rect,
            )
            return False
        pyautogui.doubleClick(x, y)
        return True
    except Exception as e:
        log.debug("try_double_click failed: %s", e)
        return False


def try_right_click(control, win_rect=None) -> bool:
    if control is None:
        return False
    try:
        r = control.BoundingRectangle
        x = int((r.left + r.right) / 2)
        y = int((r.top + r.bottom) / 2)
        if x == 0 and y == 0 and r.right == 0 and r.bottom == 0:
            return False
        if win_rect and not (win_rect[0] <= x <= win_rect[2] and
                             win_rect[1] <= y <= win_rect[3]):
            log.warning(
                "Right-click (%d,%d) is outside window bounds %s — skipping",
                x, y, win_rect,
            )
            return False
        pyautogui.rightClick(x, y)
        return True
    except Exception as e:
        log.debug("try_right_click failed: %s", e)
        return False


def type_into_focused(text: str) -> bool:
    focused = uia.GetFocusedControl()
    if focused is None:
        return False

    try:
        sv = getattr(focused, "SetValue", None)
        if callable(sv):
            sv(text)
            return True
    except Exception:
        pass

    try:
        try:
            current = focused.GetValuePattern().CurrentValue
            if current and not current[-1].isspace():
                text = " " + text
        except Exception:
            pass
        uia.SendKeys(text, waitTime=0)
        return True
    except Exception:
        return False
