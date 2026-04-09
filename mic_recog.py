"""Standalone microphone recognition test script.

Run this directly to verify VOSK speech recognition is working.
Uses the model selected in user settings (same as the main assistant).

Controls:
  ESC        — stop listening
  SHIFT + P  — print all captured results so far (debug)
"""
import json
import logging

import keyboard
import pyaudio
from vosk import Model, KaldiRecognizer

import user_settings

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("mic_recog")

MODEL_PATH = user_settings.get_model_path()
log.info("Loading model from: %s", MODEL_PATH)
model = Model(MODEL_PATH)
recognizer = KaldiRecognizer(model, 16000)

mic = pyaudio.PyAudio()
stream = mic.open(
    format=pyaudio.paInt16, channels=1, rate=16000,
    input=True, frames_per_buffer=8192,
)
stream.start_stream()

log.info("Listening. Press ESC to stop, SHIFT+P to dump all results.")

running = True
dump_results = False
all_results = []


def _stop():
    global running
    log.info("ESC pressed — stopping.")
    running = False


def _dump():
    global dump_results
    dump_results = True


keyboard.add_hotkey("esc", _stop)
keyboard.add_hotkey("shift+p", _dump)

try:
    while running:
        data = stream.read(4096, exception_on_overflow=False)

        if recognizer.AcceptWaveform(data):
            result = json.loads(recognizer.Result())
            all_results.append(result)
            text = result.get("text", "").strip()
            if text:
                log.info("Heard: %s", text)

        if dump_results:
            dump_results = False
            log.info("--- Captured results (%d) ---", len(all_results))
            for i, r in enumerate(all_results, start=1):
                log.info("  [%d] %s", i, json.dumps(r))
            log.info("--- End of results ---")

finally:
    running = False
    stream.stop_stream()
    stream.close()
    mic.terminate()
    log.info("Finished.")
