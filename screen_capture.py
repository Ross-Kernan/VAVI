import mss # https://python-mss.readthedocs.io/
import numpy as np
import cv2
import base64

class ScreenCapture:
    def __init__(self, monitor_index=1):
        self.monitor_index = monitor_index # Set which monitor to capture (1 = primary monitor)

    def capture_raw(self):
        with mss.mss() as sct: # Capture screen and return it as raw NumPy image
            monitor = sct.monitors[self.monitor_index]
            screenshot = sct.grab(monitor)
            img = np.array(screenshot) # Convert screenshot to NumPy array
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR) # Change format from BGRA to BGR (OpenCV format)
            return img

    def capture_png_bytes(self):
        img = self.capture_raw() # Capture screen and return image as PNG file bytes
        success, buffer = cv2.imencode('.png', img)
        return buffer.tobytes() if success else None

    def capture_base64(self):
        png_bytes = self.capture_png_bytes() # Capture the screen and return it as a Base64 encoded string
        return base64.b64encode(png_bytes).decode('utf-8') # Convert PNG bytes to Base64 text


if __name__ == "__main__":
    sc = ScreenCapture()
    img = sc.capture_raw()

    cv2.imshow("Screen Capture Test", img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    print("Base64 size:", len(sc.capture_base64()))

