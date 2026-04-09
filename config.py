import os
import logging

_HERE = os.path.dirname(os.path.abspath(__file__))

VOSK_MODEL_PATH = os.path.join(_HERE, "vosk-model-en-us-0.22")

TESSERACT_CMD = os.environ.get(
    "TESSERACT_CMD",
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
)

HF_TOKEN = os.environ.get("HF_TOKEN")

LISTEN_TIMEOUT = 6.0
LISTEN_SILENCE_TIMEOUT = 0.6   
ANNOUNCE_COOLDOWN = 2.0
LOOP_SLEEP = 0.05
COMMAND_SLEEP = 0.2

MAX_SCAN_DEPTH = 2000
POI_MAX_LENGTH = 3000
OCR_CONFIDENCE_THRESHOLD = 50
MAX_UIA_CHILDREN = 15

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)

log = logging.getLogger("vavi")

import uiautomation as _uia  
_uia.Logger.SetLogFile("")
