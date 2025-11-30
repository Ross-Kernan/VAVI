import pyaudio # https://people.csail.mit.edu/hubert/pyaudio/docs/
import json
import keyboard

from vosk import Model, KaldiRecognizer

MODEL_PATH = "vosk-model-small-en-us-0.15"  # Change this if needed

# Load the model
print("Loading")
model = Model(MODEL_PATH)
recognizer = KaldiRecognizer(model, 16000)

# Start microphone stream
mic = pyaudio.PyAudio()
stream = mic.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8192)
stream.start_stream()

print("Listening Now, Esc to cancel\n")

running = True
print_json = False
all_results = [] # Debug list

def stop():
    global running
    print("ESC pressed")
    running = False

def show_json(): # Debug, remove me later
    global print_json
    print_json = True

# need seperate thread for the keybinds
keyboard.add_hotkey('esc', stop)
keyboard.add_hotkey('shift+p', show_json)

#   According to boppreh the add_hotkey has its own background thread so dont need to it
#   https://github.com/boppreh/keyboard/issues/457


try:
    while running:
        data = stream.read(4096, exception_on_overflow=False)

        if recognizer.AcceptWaveform(data):
            result = json.loads(recognizer.Result())
            all_results.append(result) # Store the results in the list instead so i can see them
            text = result.get("text", "")
            if text:
                print(f"You said: {text}")

            if print_json:
                print_json = False
                print("\n---------------START--------------")
                for i, res in enumerate(all_results, start=1):
                    print(f"\nResult {i}:")
                    print(json.dumps(res, indent=4))
                print("\n--------------END--------------")
finally:
    running = False
    stream.stop_stream()
    stream.close()
    mic.terminate()
    print("Finished Up Here")
