import re
import logging
from dataclasses import dataclass
from difflib import get_close_matches
from typing import Optional

from commands.audio import _ORDINAL_MAP

log = logging.getLogger("vavi.commands.parser")

COMMAND_ALIASES = {
    "describe":    ["describe", "describe screen", "what's on screen", "points of interest"],
    "quit":        ["quit", "exit", "stop"],
    "read":        ["read", "read again", "repeat"],
    "options":     ["what can i click", "what can i press", "options"],
    "help":        ["help", "commands", "what can i say"],
    "opencv_start":["tap to start", "start game", "click start"],
    "copy":        ["copy", "copy that", "copy this"],
    "paste":       ["paste", "paste that", "paste it", "paste here"],
    "cut":         ["cut", "cut that", "cut this"],
    "select_all":  ["select all", "select everything"],
}

_ALL_PHRASES = []
_PHRASE_TO_KIND = {}
for _kind, _phrases in COMMAND_ALIASES.items():
    for _p in _phrases:
        _ALL_PHRASES.append(_p)
        _PHRASE_TO_KIND[_p] = _kind


@dataclass
class Command:
    kind: str
    arg: Optional[str] = None


def parse_command(text: str) -> Optional[Command]:
    if not text:
        return None

    text = text.strip().lower()

    for kind, phrases in COMMAND_ALIASES.items():
        if text in phrases:
            return Command(kind)

    m = re.match(r"^click\s+(?:the\s+)?(\w+)\s+result$", text)
    if m:
        tok = m.group(1).strip()
        _n = _ORDINAL_MAP.get(tok)
        if _n is None:
            try:
                _n = int(tok)
            except ValueError:
                pass
        if _n is not None:
            return Command("result", str(_n))

    m = re.match(r"^click\s+result\s+(\w+)$", text)
    if m:
        tok = m.group(1).strip()
        _n = _ORDINAL_MAP.get(tok)
        if _n is None:
            try:
                _n = int(tok)
            except ValueError:
                pass
        if _n is not None:
            return Command("result", str(_n))

    m = re.match(r"^double\s+(?:click|tap)\s+(.+)$", text)
    if m:
        return Command("double_click", m.group(1).strip())

    m = re.match(r"^right\s+(?:click|tap)\s+(.+)$", text)
    if m:
        return Command("right_click", m.group(1).strip())

    m = re.match(r"^click\s+(.+)$", text)
    if m:
        return Command("click", m.group(1).strip())

    m = re.match(r"^type\s+(.+)$", text)
    if m:
        return Command("type", m.group(1).strip())

    m = re.match(r"^spell\s+(.+)$", text)
    if m:
        return Command("spell", m.group(1).strip())

    m = re.match(r"^press\s+(.+)$", text)
    if m:
        return Command("press", m.group(1).strip())

    if re.match(r"^(go\s+)?back(\s+page)?$", text) or text == "page back":
        return Command("back")

    m = re.match(r"^open\s+(.+)$", text)
    if m:
        arg = m.group(1).strip()
        if re.match(r"^(new|close|next|previous|prev|last)\s+tab$", arg):
            return Command("tab", arg.split()[0])
        return Command("open", arg)

    m = re.match(r"^launch\s+(.+)$", text)
    if m:
        return Command("open", m.group(1).strip())

    m = re.match(r"^(?:set\s+)?volume\s+(?:to\s+)?(.+)$", text)
    if m:
        return Command("volume", m.group(1).strip())

    m = re.match(r"^scroll\s+(up|down)(?:\s+(.+))?$", text)
    if m:
        direction = m.group(1)
        amount_str = (m.group(2) or "").strip()
        return Command("scroll", f"{direction} {amount_str}".strip())

    m = re.match(r"^search\s+(.+)$", text)
    if m:
        return Command("search", m.group(1).strip())

    m = re.match(r"^(new|close|next|previous|prev|last)\s+tab$", text)
    if m:
        return Command("tab", m.group(1).strip())

    m = re.match(r"^close\s+tab\s+(.+)$", text)
    if m:
        return Command("tab", f"close {m.group(1).strip()}")

    m = re.match(r"^switch\s+(?:to\s+)?tab\s+(.+)$", text)
    if m:
        return Command("tab", m.group(1).strip())

    m = re.match(r"^tab\s+(.+)$", text)
    if m:
        return Command("tab", m.group(1).strip())

    m = re.match(r"^(pause|play|mute|unmute)$", text)
    if m:
        return Command("media", m.group(1).strip())

    m = re.match(r"^(minimize|minimise|maximize|maximise|restore|close)(?:\s+\S+)?$", text)
    if m:
        return Command("click", m.group(1))

    m = re.match(r"^describe\s+(.+)$", text)
    if m:
        return Command("describe", m.group(1).strip())

    close = get_close_matches(text, _ALL_PHRASES, n=1, cutoff=0.7) if len(text) >= 4 else []
    if close:
        matched_phrase = close[0]
        kind = _PHRASE_TO_KIND[matched_phrase]
        log.info("Fuzzy matched '%s' -> '%s' (%s)", text, matched_phrase, kind)
        return Command(kind)

    return Command("unknown", text)


HELP_TEXT = (
    "Available commands: "
    "Describe, to hear what's on screen. "
    "Describe, followed by a question, to ask something specific about the screen. "
    "Click, followed by a name. "
    "Double click, followed by a name, to double-click it. Use this to open files and folders in File Explorer. "
    "Open, followed by an app name or website to launch it. If a browser is already open, websites open in a new tab. "
    "New tab, to open a new browser tab. "
    "Close tab, to close the current tab. Close tab, followed by a number, to close a specific tab. "
    "Next tab or previous tab, to switch tabs. "
    "Switch to tab, followed by a number, or just tab followed by a number, to jump to a specific tab. "
    "Pause or play, to toggle media playback. Mute or unmute, to toggle audio. "
    "Search, followed by a search term, to find the right search bar, type your query, and press enter. "
    "Click the first result, or click result followed by a number, to click a numbered search result in a browser. "
    "For example: click the second result, or click result three. "
    "Type, followed by text. "
    "Spell, followed by letters, to type them joined together without spaces. For example, spell p g e types pge. "
    "Right click, followed by a name, to right-click it. "
    "Copy, paste, or cut, to use the clipboard. Select all, to select everything. "
    "Press, followed by a key name, to press it. For example: press escape, press f5, press ctrl z, press alt f4. "
    "Set volume to, followed by a number from zero to one hundred. "
    "Scroll up or scroll down, optionally followed by a number. Add the word lines for fine-grained line scrolling. "
    "Read, to repeat. "
    "What can I click, for options. "
    "Help, for this list. "
    "Quit, to stop."
)
