import queue
import threading
import time
import pythoncom
import win32com.client
import logging

log = logging.getLogger("vavi.tts")

_SVSFlagsAsync = 1
_SVSFPurgeBeforeSpeak = 2


class TextToSpeech:
    def __init__(self, rate=0, volume=100, voice_contains=None):
        self.q = queue.Queue()
        self.rate = rate
        self.volume = volume
        self.voice_contains = (voice_contains or "").lower()

        self._stop_event = threading.Event()
        self._interrupt_event = threading.Event()
        self._speaking = threading.Event()
        self._voice = None          

        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def _run(self):
        pythoncom.CoInitialize()
        try:
            voice = win32com.client.Dispatch("SAPI.SpVoice")
            voice.Rate = self.rate
            voice.Volume = self.volume

            if self.voice_contains:
                try:
                    voices = voice.GetVoices()
                    for i in range(voices.Count):
                        v = voices.Item(i)
                        desc = v.GetDescription().lower()
                        if self.voice_contains in desc:
                            voice.Voice = v
                            break
                except Exception:
                    pass

            self._voice = voice

            while not self._stop_event.is_set():
                if self._interrupt_event.is_set():
                    self._interrupt_event.clear()
                    try:
                        voice.Speak("", _SVSFPurgeBeforeSpeak)
                    except Exception:
                        pass
                    self._speaking.clear()
                    continue

                try:
                    text = self.q.get(timeout=0.1)
                except queue.Empty:
                    continue

                if text is None:
                    break

                if self._interrupt_event.is_set():
                    self._interrupt_event.clear()
                    try:
                        voice.Speak("", _SVSFPurgeBeforeSpeak)
                    except Exception:
                        pass
                    self._speaking.clear()
                    continue

                try:
                    log.info("Speaking: %s", text)
                    self._speaking.set()
                    voice.Speak(text, _SVSFlagsAsync)
                    while voice.Status.RunningState != 1:
                        if self._interrupt_event.is_set():
                            self._interrupt_event.clear()
                            voice.Speak("", _SVSFPurgeBeforeSpeak)
                            break
                        time.sleep(0.05)
                    self._speaking.clear()
                except Exception as e:
                    log.error("TTS error: %s", e)
                    self._speaking.clear()

        finally:
            pythoncom.CoUninitialize()

    @property
    def speaking(self) -> bool:
        return self._speaking.is_set()

    def speak(self, text: str):
        if text:
            self.q.put(text)

    def interrupt(self):
        while not self.q.empty():
            try:
                self.q.get_nowait()
            except queue.Empty:
                break
        if self._speaking.is_set():
            self._interrupt_event.set()

    def close(self):
        self._stop_event.set()
        self.q.put(None)
        self.thread.join(timeout=3.0)


if __name__ == "__main__":
    tts = TextToSpeech()
    tts.speak("Hello, this is a test.")
    import time
    time.sleep(3)
    tts.close()
