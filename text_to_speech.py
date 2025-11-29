import pyttsx3 # https://pyttsx3.readthedocs.io/en/latest/

# Dont Know if pyttsx3 runs on its own thread
# may need to be implemented later to not block microphone input

class TextToSpeech:
    def __init__(self):
        """Initialize the TTS engine"""
        self.engine = pyttsx3.init()
        
        # Configure voice speed and volume
        self.engine.setProperty('rate', 200) # Between 100 - 200
        self.engine.setProperty('volume', 0.4) # increase later, .9 was a sweet spot for actual usage
        
        voices = self.engine.getProperty('voices')
        #self.engine.setProperty('voice', voices[0].id)  # Default voice
        self.engine.setProperty('voice', voices[1].id)  # Softer Female voice
    
    def speak(self, text):
        print(f"Speaking: {text}")
        self.engine.say(text)
        self.engine.runAndWait()
        self.engine.stop()
    
    def stop(self):
        """Stop current speech"""
        self.engine.stop()


if __name__ == "__main__":
    tts = TextToSpeech()
    
    message = "The quick brown fox jumps over the lazy dog"
    tts.speak(message)

    # --- Test List delete me later ---
    # Hello, this is a test of the text to speech system 
    # A... e... i... o... u...
    # Brown, Zebra, Motorway, Brake, Moon, Juice, Krugs
    # The quick brown fox jumps over the lazy dog
    # --- Test List delete me later ---