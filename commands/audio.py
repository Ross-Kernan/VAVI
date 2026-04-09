import logging
from typing import Optional

log = logging.getLogger("vavi.commands.audio")

_SPOKEN_NUMBERS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
    "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17,
    "eighteen": 18, "nineteen": 19, "twenty": 20, "thirty": 30,
    "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
    "eighty": 80, "ninety": 90, "hundred": 100,
}

_ORDINAL_MAP = {
    "first": 1,  "1st": 1,  "one": 1,   "1": 1,
    "second": 2, "2nd": 2,  "two": 2,   "2": 2,
    "third": 3,  "3rd": 3,  "three": 3, "3": 3,
    "fourth": 4, "4th": 4,  "four": 4,  "4": 4,
    "fifth": 5,  "5th": 5,  "five": 5,  "5": 5,
    "sixth": 6,  "6th": 6,  "six": 6,   "6": 6,
    "seventh": 7,"7th": 7,  "seven": 7, "7": 7,
    "eighth": 8, "8th": 8,  "eight": 8, "8": 8,
    "ninth": 9,  "9th": 9,  "nine": 9,  "9": 9,
    "tenth": 10, "10th": 10,"ten": 10,  "10": 10,
}


def _words_to_digits(text: str) -> str:
    # Replace spoken number words in *text* with their digit equivalents
    tokens = text.split()
    result = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok in _SPOKEN_NUMBERS:
            val = _SPOKEN_NUMBERS[tok]
            if (i + 1 < len(tokens)
                    and tokens[i + 1] in _SPOKEN_NUMBERS
                    and val >= 20
                    and 1 <= _SPOKEN_NUMBERS[tokens[i + 1]] <= 9):
                result.append(str(val + _SPOKEN_NUMBERS[tokens[i + 1]]))
                i += 2
                continue
            result.append(str(val))
        else:
            result.append(tok)
        i += 1
    return " ".join(result)


def _parse_volume_level(s: str) -> Optional[int]:
    # Convert a spoken volume string to an integer 0-100
    s = s.strip().lower()

    try:
        return max(0, min(100, int(s)))
    except ValueError:
        pass

    if s in _SPOKEN_NUMBERS:
        return _SPOKEN_NUMBERS[s]

    parts = s.split()
    if len(parts) == 2 and parts[0] in _SPOKEN_NUMBERS and parts[1] in _SPOKEN_NUMBERS:
        return max(0, min(100, _SPOKEN_NUMBERS[parts[0]] + _SPOKEN_NUMBERS[parts[1]]))

    return None


def set_system_volume(level: int) -> bool:
    # Set the Windows master volume to level (0-100). Returns True on success.
    level = max(0, min(100, level))
    try:
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        from comtypes import CLSCTX_ALL
        from ctypes import cast, POINTER

        devices = AudioUtilities.GetSpeakers()
        interface = devices._dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level / 100.0, None)
        log.info("System volume set to %d%%", level)
        return True
    except Exception as e:
        log.error("set_system_volume failed: %s", e)
        return False
