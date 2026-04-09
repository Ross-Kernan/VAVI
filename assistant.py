import time
import json
import logging
import threading
from typing import Optional, Tuple

import os

import keyboard
import pyautogui
import win32gui
import win32process

import config
import user_settings
from config import (
    VOSK_MODEL_PATH,
    LISTEN_TIMEOUT,
    LISTEN_SILENCE_TIMEOUT,
    ANNOUNCE_COOLDOWN,
    MAX_SCAN_DEPTH,
    POI_MAX_LENGTH,
    LOOP_SLEEP,
    COMMAND_SLEEP,
)
from desktop_probe import capture_desktop_state
from text_to_speech import TextToSpeech
from opencv_click import find_button_regions, click_region
from screen_capture import ScreenCapture

import pyaudio
from vosk import Model, KaldiRecognizer

import uiautomation as uia

from poi_llm import points_of_interest

from commands.uia_utils import (
    describe_focus, normalize_window_title, summarize_actionables,
    iter_descendants, find_control_by_name,
    try_invoke_or_click, try_double_click, try_right_click, type_into_focused,
    _CLICK_TARGET_ALIASES, _INPUT_SUFFIX_RE, _INPUT_PREFERRED_TYPES,
)
from commands.input_actions import try_window_command, press_key
from commands.app_launcher import launch_application, _WEBSITE_ALIASES, _click_taskbar_app
from commands.audio import _words_to_digits, _parse_volume_level, set_system_volume
from commands.scroll import _focus_target_window, scroll_page, scroll_lines
from commands.browser import (
    click_text_via_ocr, click_browser_result, perform_search, handle_tab_command,
)
from commands.parser import Command, COMMAND_ALIASES, parse_command, HELP_TEXT

log = logging.getLogger("vavi.assistant")

stop_event = threading.Event()

class SpeechRecognizer:
    def __init__(self, model_path: str = VOSK_MODEL_PATH, rate: int = 16000):
        self.model = Model(model_path)
        self.recognizer = KaldiRecognizer(self.model, rate)

        self.rate = rate
        self.pa = pyaudio.PyAudio()
        self.stream = self.pa.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=rate,
            input=True,
            frames_per_buffer=8192,
        )
        self.stream.start_stream()

    def _flush_audio_buffer(self):
        try:
            available = self.stream.get_read_available()
            while available > 0:
                self.stream.read(min(available, 4096), exception_on_overflow=False)
                available = self.stream.get_read_available()
        except Exception:
            pass
        self.recognizer.Reset()

    def listen_for_phrase(self, max_seconds: float = LISTEN_TIMEOUT) -> Optional[str]:
        self._flush_audio_buffer()
        deadline = time.time() + max_seconds
        chunks = []
        last_chunk_time: Optional[float] = None

        while time.time() < deadline:
            # Once speech has started, stop when silence window expires
            if last_chunk_time is not None:
                if (time.time() - last_chunk_time) >= LISTEN_SILENCE_TIMEOUT:
                    break

            data = self.stream.read(4096, exception_on_overflow=False)
            if self.recognizer.AcceptWaveform(data):
                result = json.loads(self.recognizer.Result())
                text = (result.get("text") or "").strip().lower()
                if text:
                    chunks.append(text)
                    last_chunk_time = time.time()

        if chunks:
            return " ".join(chunks)

        partial = json.loads(self.recognizer.PartialResult()).get("partial", "").strip().lower()
        return partial if partial else None

    def close(self):
        try:
            self.stream.stop_stream()
            self.stream.close()
        finally:
            self.pa.terminate()


PUSH_TO_TALK_KEYS = [
    "right ctrl",
    "left ctrl",
]

def is_push_to_talk_pressed() -> bool:
    return any(keyboard.is_pressed(k) for k in PUSH_TO_TALK_KEYS)


class _StateCache:
    def __init__(self, max_age: float = 5.0):
        self.max_age = max_age
        self.state: Optional[dict] = None
        self.signature: Optional[tuple] = None
        self.timestamp: float = 0.0

    def get(self, force_ocr: bool = False, target_hwnd=None) -> dict:
        """Return a (possibly cached) desktop state."""
        now = time.time()
        need_refresh = (
            self.state is None
            or (now - self.timestamp) > self.max_age
        )

        if need_refresh:
            try:
                self.state = capture_desktop_state(
                    do_ocr=force_ocr, target_hwnd=target_hwnd)
            except Exception as e:
                log.error("capture_desktop_state failed: %s", e)
                if self.state is None:
                    self.state = {"window": {"title": "Unknown", "bounds": {}}, "uia": {}}
                return self.state
            self.timestamp = now
            w = ((self.state.get("window") or {}).get("title") or "").strip()
            u = self.state.get("uia") or {}
            f = u.get("focused") or {}
            self.signature = (
                normalize_window_title(w),
                (f.get("name") or "").strip(),
                (f.get("type") or "").strip(),
            )
        elif force_ocr and "ocr_text" not in (self.state or {}):
            try:
                self.state = capture_desktop_state(
                    do_ocr=True, target_hwnd=target_hwnd)
                self.timestamp = now
            except Exception as e:
                log.error("OCR re-probe failed: %s", e)

        return self.state

    def invalidate(self):
        self.state = None

def main():
    tts = TextToSpeech()
    sr = SpeechRecognizer(model_path=user_settings.get_model_path())

    _active_ptt_keys = user_settings.get_ptt_keys()
    log.info("Push-to-talk keys: %s", _active_ptt_keys)

    _my_pid = os.getpid()
    target_hwnd: Optional[int] = None  

    cache = _StateCache()
    last_signature: Optional[Tuple[str, str, str]] = None
    last_announce_time = 0.0

    last_poi_signature: Optional[Tuple[str, str, str]] = None
    last_poi_text: Optional[str] = None

    ptt_was_down = False

    tts.speak(
        "Assistant started. Hold control and speak a command. "
        "Hold control and say help for a list of commands."
    )

    stop_event.clear()
    try:
        while not stop_event.is_set():
            current_fg = win32gui.GetForegroundWindow()
            if current_fg:
                try:
                    _, _fg_pid = win32process.GetWindowThreadProcessId(current_fg)
                except Exception:
                    _fg_pid = _my_pid  
                if _fg_pid != _my_pid:
                    target_hwnd = current_fg

            if target_hwnd and not win32gui.IsWindow(target_hwnd):
                log.debug("Clearing stale target_hwnd %s", target_hwnd)
                target_hwnd = None

            state = cache.get(force_ocr=False, target_hwnd=target_hwnd)

            if "uia_error" in state:
                log.warning("UIA error in state: %s", state["uia_error"])
            if "ocr_error" in state:
                log.warning("OCR error in state: %s", state["ocr_error"])

            signature = cache.signature or ("", "", "")
            now = time.time()
            if signature != last_signature and (now - last_announce_time) > ANNOUNCE_COOLDOWN:
                tts.speak(describe_focus(state))
                last_signature = signature
                last_announce_time = now

            ptt_down = any(keyboard.is_pressed(k) for k in _active_ptt_keys)
            if not ptt_down:
                ptt_was_down = False
                time.sleep(LOOP_SLEEP)
                continue

            if ptt_was_down:
                time.sleep(LOOP_SLEEP)
                continue

            ptt_was_down = True
            log.info("Push-to-talk active. Listening...")

            cmd_text = sr.listen_for_phrase()
            if not cmd_text:
                continue

            log.info("Heard: %s", cmd_text)

            cmd = parse_command(cmd_text)
            if cmd is None:
                continue

            tts.interrupt()

            cache.invalidate()

            if cmd.kind == "quit":
                tts.speak("Are you sure you want to quit? Say yes to confirm or no to cancel.")
                time.sleep(0.15)
                while tts.speaking:
                    time.sleep(0.05)
                confirm_text = sr.listen_for_phrase()
                if confirm_text and "yes" in confirm_text.lower():
                    tts.speak("Stopping.")
                    break
                else:
                    tts.speak("Cancelled.")

            elif cmd.kind == "help":
                tts.speak(HELP_TEXT)

            elif cmd.kind == "read":
                state = cache.get(force_ocr=False, target_hwnd=target_hwnd)
                tts.speak(describe_focus(state))

            elif cmd.kind == "options":
                state = cache.get(force_ocr=False, target_hwnd=target_hwnd)
                tts.speak(summarize_actionables(state))

            elif cmd.kind == "click":
                target = cmd.arg or ""

                if try_window_command(target, hwnd=target_hwnd):
                    tts.speak(f"{target.capitalize()}d.")
                    time.sleep(COMMAND_SLEEP)
                    continue

                target_alt = _words_to_digits(target)

                state = cache.get(force_ocr=True, target_hwnd=target_hwnd)
                control = find_control_by_name(
                    target, max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)
                if control is None and target_alt != target:
                    control = find_control_by_name(
                        target_alt, max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)

                win_rect = None
                if target_hwnd:
                    try:
                        win_rect = win32gui.GetWindowRect(target_hwnd)
                    except Exception:
                        pass

                if control is None:
                    ok = click_text_via_ocr(state, target)
                    if not ok and target_alt != target:
                        ok = click_text_via_ocr(state, target_alt)
                    if ok:
                        cache.invalidate()
                        tts.speak(f"Clicked {target}.")
                    else:
                        tts.speak(f"I couldn't find {target}.")
                else:
                    try:
                        ok = try_invoke_or_click(control, win_rect=win_rect)
                    except Exception as e:
                        log.error("UIA click failed: %s", e)
                        ok = False

                    if not ok:
                        ok = click_text_via_ocr(state, target)
                        if not ok and target_alt != target:
                            ok = click_text_via_ocr(state, target_alt)

                    try:
                        spoken = (control.Name or "").strip() or target
                    except Exception:
                        spoken = target

                    if ok:
                        cache.invalidate()
                        tts.speak(f"Clicked {spoken}.")
                    else:
                        tts.speak(f"I found {spoken}, but couldn't click it.")

            elif cmd.kind == "double_click":
                target = cmd.arg or ""

                target_alt = _words_to_digits(target)

                state = cache.get(force_ocr=True, target_hwnd=target_hwnd)
                control = find_control_by_name(
                    target, max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)
                if control is None and target_alt != target:
                    control = find_control_by_name(
                        target_alt, max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)

                win_rect = None
                if target_hwnd:
                    try:
                        win_rect = win32gui.GetWindowRect(target_hwnd)
                    except Exception:
                        pass

                if control is None:
                    ok = click_text_via_ocr(state, target, double=True)
                    if not ok and target_alt != target:
                        ok = click_text_via_ocr(state, target_alt, double=True)
                    if ok:
                        cache.invalidate()
                        tts.speak(f"Double-clicked {target}.")
                    else:
                        tts.speak(f"I couldn't find {target}.")
                else:
                    try:
                        ok = try_double_click(control, win_rect=win_rect)
                    except Exception as e:
                        log.error("UIA double-click failed: %s", e)
                        ok = False

                    if not ok:
                        ok = click_text_via_ocr(state, target, double=True)
                        if not ok and target_alt != target:
                            ok = click_text_via_ocr(state, target_alt, double=True)

                    try:
                        spoken = (control.Name or "").strip() or target
                    except Exception:
                        spoken = target

                    if ok:
                        cache.invalidate()
                        tts.speak(f"Double-clicked {spoken}.")
                    else:
                        tts.speak(f"I found {spoken}, but couldn't double-click it.")

            elif cmd.kind == "right_click":
                target = cmd.arg or ""
                target_alt = _words_to_digits(target)

                state = cache.get(force_ocr=True, target_hwnd=target_hwnd)
                control = find_control_by_name(
                    target, max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)
                if control is None and target_alt != target:
                    control = find_control_by_name(
                        target_alt, max_nodes=MAX_SCAN_DEPTH, hwnd=target_hwnd)

                win_rect = None
                if target_hwnd:
                    try:
                        win_rect = win32gui.GetWindowRect(target_hwnd)
                    except Exception:
                        pass

                if control is None:
                    ok = click_text_via_ocr(state, target, right=True)
                    if not ok and target_alt != target:
                        ok = click_text_via_ocr(state, target_alt, right=True)
                    if ok:
                        cache.invalidate()
                        tts.speak(f"Right-clicked {target}.")
                    else:
                        tts.speak(f"I couldn't find {target}.")
                else:
                    try:
                        ok = try_right_click(control, win_rect=win_rect)
                    except Exception as e:
                        log.error("Right-click failed: %s", e)
                        ok = False

                    if not ok:
                        ok = click_text_via_ocr(state, target, right=True)
                        if not ok and target_alt != target:
                            ok = click_text_via_ocr(state, target_alt, right=True)

                    try:
                        spoken = (control.Name or "").strip() or target
                    except Exception:
                        spoken = target

                    if ok:
                        cache.invalidate()
                        tts.speak(f"Right-clicked {spoken}.")
                    else:
                        tts.speak(f"I found {spoken}, but couldn't right-click it.")

            elif cmd.kind == "type":
                text = cmd.arg or ""
                if not text:
                    tts.speak("What should I type?")
                else:
                    ok = type_into_focused(text)
                    tts.speak("Typed." if ok else "I couldn't type into the focused control.")

            elif cmd.kind == "spell":
                text = cmd.arg or ""
                if not text:
                    tts.speak("What should I spell?")
                else:
                    spelled = text.replace(" ", "")
                    ok = type_into_focused(spelled)
                    tts.speak("Typed." if ok else "I couldn't type into the focused control.")

            elif cmd.kind == "copy":
                _focus_target_window(target_hwnd)
                pyautogui.hotkey("ctrl", "c")
                tts.speak("Copied.")

            elif cmd.kind == "paste":
                _focus_target_window(target_hwnd)
                pyautogui.hotkey("ctrl", "v")
                tts.speak("Pasted.")

            elif cmd.kind == "cut":
                _focus_target_window(target_hwnd)
                pyautogui.hotkey("ctrl", "x")
                tts.speak("Cut.")

            elif cmd.kind == "select_all":
                _focus_target_window(target_hwnd)
                pyautogui.hotkey("ctrl", "a")
                tts.speak("Selected all.")

            elif cmd.kind == "press":
                key = cmd.arg or ""
                _focus_target_window(target_hwnd)
                ok = press_key(key)
                tts.speak(f"Pressed {key}." if ok else f"I can't press {key}.")

            elif cmd.kind == "opencv_start":
                tts.speak("Looking for a start button.")

                import mss as _mss
                monitor_index = 1
                sc = ScreenCapture(monitor_index=monitor_index)
                img = sc.capture_raw()

                with _mss.mss() as sct:
                    mon = sct.monitors[monitor_index]
                    mon_offset = (mon["left"], mon["top"])

                regions = find_button_regions(img)

                if not regions:
                    tts.speak("I could not find a start button.")
                else:
                    target = regions[0]
                    tts.speak("Clicking start.")
                    click_region(target, offset=mon_offset)

            elif cmd.kind == "open":
                app_name = cmd.arg or ""
                if not app_name:
                    tts.speak("What should I open?")
                else:
                    state = cache.get(force_ocr=False, target_hwnd=target_hwnd)
                    window_title = ((state.get("window") or {}).get("title") or "").lower()
                    in_browser = any(
                        b in window_title
                        for b in ("chrome", "firefox", "edge", "mozilla")
                    )
                    url = _WEBSITE_ALIASES.get(app_name.strip().lower())

                    if url and in_browser:
                        tts.speak(f"Opening {app_name} in a new tab.")
                        try:
                            _focus_target_window(target_hwnd)
                            pyautogui.hotkey("ctrl", "t")
                            time.sleep(0.5)            
                            uia.SendKeys(url, waitTime=0)
                            time.sleep(0.1)
                            uia.SendKeys("{Enter}", waitTime=0)
                            cache.invalidate()
                        except Exception as e:
                            log.warning("New-tab open failed: %s", e)
                            tts.speak(f"I couldn't open a new tab for {app_name}.")
                    else:
                        tts.speak(f"Opening {app_name}.")
                        ok = launch_application(app_name)
                        if not ok:
                            tts.speak(f"I couldn't open {app_name}.")

            elif cmd.kind == "tab":
                _focus_target_window(target_hwnd)
                ok, message = handle_tab_command(cmd.arg or "")
                tts.speak(message)
                if ok:
                    cache.invalidate()

            elif cmd.kind == "search":
                query = cmd.arg or ""
                if not query:
                    tts.speak("What would you like to search for?")
                else:
                    _focus_target_window(target_hwnd)
                    state = cache.get(force_ocr=False, target_hwnd=target_hwnd)
                    ok, message = perform_search(query, state, target_hwnd=target_hwnd)
                    tts.speak(message)
                    if ok:
                        cache.invalidate()

            elif cmd.kind == "media":
                action = cmd.arg or ""
                _focus_target_window(target_hwnd)
                if action in ("pause", "play"):
                    pyautogui.press("k")
                    tts.speak(f"{action.capitalize()}d.")
                elif action == "mute":
                    uia.SendKeys("m", waitTime=0)
                    tts.speak("Muted.")
                elif action == "unmute":
                    uia.SendKeys("m", waitTime=0)
                    tts.speak("Unmuted.")

            elif cmd.kind == "volume":
                raw = cmd.arg or ""
                level = _parse_volume_level(raw)
                if level is None:
                    tts.speak(f"I didn't understand the volume level: {raw}.")
                else:
                    ok = set_system_volume(level)
                    if ok:
                        tts.speak(f"Volume set to {level} percent.")
                    else:
                        tts.speak("I couldn't change the volume.")

            elif cmd.kind == "scroll":
                parts = (cmd.arg or "down").split(None, 1)
                direction = parts[0] if parts[0] in ("up", "down") else "down"
                amount_str = parts[1] if len(parts) > 1 else ""

                def _parse_amount(s):
                    words = s.split()
                    for n in range(len(words), 0, -1):
                        v = _parse_volume_level(" ".join(words[:n]))
                        if v is not None:
                            return v
                    return None

                if "line" in amount_str:
                    num_str = amount_str.replace("lines", "").replace("line", "").strip()
                    lines = _parse_amount(num_str) if num_str else None
                    lines = max(1, min(50, lines)) if lines is not None else 3
                    ok = scroll_lines(direction, lines=lines, hwnd=target_hwnd)
                else:
                    clicks = _parse_amount(amount_str) if amount_str else None
                    clicks = max(1, min(50, clicks)) if clicks is not None else 5
                    ok = scroll_page(direction, clicks=clicks, hwnd=target_hwnd)

                tts.speak(f"Scrolling {direction}." if ok else "I couldn't scroll.")

            elif cmd.kind == "result":
                result_n = int(cmd.arg or "1")
                _focus_target_window(target_hwnd)
                ok, message = click_browser_result(result_n, target_hwnd=target_hwnd)
                tts.speak(message)
                if ok:
                    cache.invalidate()

            elif cmd.kind == "back":
                _focus_target_window(target_hwnd)
                try:
                    pyautogui.hotkey("alt", "left")
                    tts.speak("Going back.")
                    cache.invalidate()
                except Exception as e:
                    log.error("back page failed: %s", e)
                    tts.speak("I couldn't go back.")

            elif cmd.kind == "describe":
                user_request = cmd.arg or None
                poi_sig = cache.signature or ("", "", "")

                if user_request is None and poi_sig == last_poi_signature and last_poi_text:
                    tts.speak(last_poi_text)
                else:
                    if user_request:
                        tts.speak("One moment.")
                    else:
                        tts.speak("Describing screen, one moment.")

                    state = cache.get(force_ocr=True, target_hwnd=target_hwnd)

                    try:
                        sc = ScreenCapture(monitor_index=1)
                        screenshot_b64 = sc.capture_base64()
                    except Exception as e:
                        log.warning("Screenshot capture failed: %s", e)
                        screenshot_b64 = None

                    try:
                        poi = points_of_interest(
                            state,
                            screenshot_b64=screenshot_b64,
                            user_request=user_request,
                        )
                    except Exception as e:
                        log.error("POI failed, falling back: %s", e)
                        poi = describe_focus(state) + " " + summarize_actionables(state)

                    poi = (poi or "").strip()
                    if not poi:
                        title = ((state.get("window") or {}).get("title") or "").strip()
                        if title:
                            poi = f"You are in {title}. " + summarize_actionables(state)
                        else:
                            poi = "I couldn't generate a description."
                    elif len(poi) > POI_MAX_LENGTH:
                        poi = poi[:POI_MAX_LENGTH] + "..."

                    if user_request is None:
                        last_poi_signature = poi_sig
                        last_poi_text = poi

                    tts.speak(poi)

            else:
                tts.speak(
                    "I didn't understand. Say help for a list of commands."
                )

            time.sleep(COMMAND_SLEEP)

    finally:
        try:
            sr.close()
        except Exception:
            pass
        try:
            tts.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
