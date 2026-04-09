import cv2
import numpy as np
import pyautogui
import logging

from screen_capture import ScreenCapture

log = logging.getLogger("vavi.opencv")

_MIN_AREA = 2400
_MAX_AREA = 120_000


def _fill_uniformity(img_region) -> float:
    if img_region.size == 0:
        return 0.0
    hsv = cv2.cvtColor(img_region, cv2.COLOR_BGR2HSV)
    std_s = np.std(hsv[:, :, 1])
    std_v = np.std(hsv[:, :, 2])
    avg_std = (std_s + std_v) / 2.0
    return max(0.0, 1.0 - avg_std / 80.0)

def _rectangularity(contour, bbox_area) -> float:
    if bbox_area <= 0:
        return 0.0
    return cv2.contourArea(contour) / bbox_area

def _confidence(contour, img, x, y, w, h) -> float:
    area = w * h
    rect_score = _rectangularity(contour, area)
    fill_score = _fill_uniformity(img[y:y+h, x:x+w])
    size_score = min(area / 10_000, 1.0)
    return 0.4 * rect_score + 0.35 * fill_score + 0.25 * size_score

def find_button_regions(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    edges = cv2.Canny(blur, 50, 150)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    regions = []
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        area = w * h

        if area < _MIN_AREA or area > _MAX_AREA:
            continue

        aspect = w / max(h, 1)
        if not (0.5 <= aspect <= 5.0):
            continue

        conf = _confidence(c, img, x, y, w, h)
        regions.append((x, y, w, h, conf))

    regions.sort(key=lambda r: r[4], reverse=True)
    log.info("Found %d button-like regions", len(regions))
    return regions


def click_region(region, offset=(0, 0)):
    x, y, w, h = region[0], region[1], region[2], region[3]
    cx = offset[0] + x + w // 2
    cy = offset[1] + y + h // 2
    pyautogui.click(cx, cy)


if __name__ == "__main__":
    sc = ScreenCapture(monitor_index=1)
    img = sc.capture_raw()

    regions = find_button_regions(img)
    print("Found regions:", len(regions))

    if regions:
        target = regions[0]
        print("Clicking:", target)
        click_region(target)
    else:
        print("No button-like regions found.")
