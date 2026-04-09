import json
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH = os.path.join(_HERE, "settings.json")

MODELS = {
    "small": "vosk-model-small-en-us-0.15",
    "full":  "vosk-model-en-us-0.22",
}

PTT_KEY_OPTIONS = {
    "ctrl":       ("Either Ctrl",  ["left ctrl", "right ctrl"]),
    "left ctrl":  ("Left Ctrl",    ["left ctrl"]),
    "right ctrl": ("Right Ctrl",   ["right ctrl"]),
    "left alt":   ("Left Alt",     ["left alt"]),
    "right alt":  ("Right Alt",    ["right alt"]),
    "caps lock":  ("Caps Lock",    ["caps lock"]),
}

DEFAULTS = {"model": "full", "ptt_key": "ctrl"}


def get_ptt_keys() -> list:
    s = load()
    ptt_setting = s.get("ptt_key", "ctrl")
    _, keys = PTT_KEY_OPTIONS.get(ptt_setting, PTT_KEY_OPTIONS["ctrl"])
    return keys


def load() -> dict:
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as fh:
            return {**DEFAULTS, **json.load(fh)}
    except (FileNotFoundError, json.JSONDecodeError):
        return dict(DEFAULTS)


def save(settings: dict) -> None:
    tmp = SETTINGS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(settings, fh, indent=2)
    os.replace(tmp, SETTINGS_PATH)


def get_model_path() -> str:
    s = load()
    model_dir = MODELS.get(s.get("model", "full"), MODELS["full"])
    return os.path.join(_HERE, model_dir)
