import cv2
import numpy as np
import pyautogui
from cvzone.HandTrackingModule import HandDetector
import time
import math
from collections import deque

# Initialize
cap = cv2.VideoCapture(0)
detector = HandDetector(detectionCon=0.8, maxHands=1)

# For motion trace ('J' and 'Z')
trail = deque(maxlen=20)
motion_triggered = False
motion_cooldown = 30
motion_timer = 0

def distance(p1, p2):
    return math.hypot(p2[0] - p1[0], p2[1] - p1[1])

# For volume control
vol_start_distance = None
vol_min = 0
vol_max = 100

while True:
    success, img = cap.read()
    if not success:
        break

    hands, img = detector.findHands(img)

    if hands:
        hand = hands[0]
        lmList = hand["lmList"]
        fingers = detector.fingersUp(hand)

        # Positions
        thumb_tip = lmList[4][:2]
        index_tip = lmList[8][:2]
        middle_tip = lmList[12][:2]
        pinky_tip = lmList[20][:2]

        # ================= Pinch Gestures =================
        # Left Click Gesture → Next Slide
        if distance(thumb_tip, index_tip) < 40:
            pyautogui.press("right")  # Move to next slide
            cv2.putText(img, "Next Slide (Right Arrow)", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
            time.sleep(0.3)

        # Right Click Gesture → Previous Slide
        elif distance(thumb_tip, pinky_tip) < 40:
            pyautogui.press("left")  # Move to previous slide
            cv2.putText(img, "Previous Slide (Left Arrow)", (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
            time.sleep(0.3)

        # Volume Control: Thumb + Middle
        current_time = time.time()

        if distance(thumb_tip, middle_tip) < 40:
            if vol_start_distance is None:
                vol_start_distance = current_time  # Start timer
            elif current_time - vol_start_distance >= 1.5:
                pyautogui.press("volumedown")
                cv2.putText(img, "Volume Down", (10, 100), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 100, 0), 2)
                vol_start_distance = current_time  # Reset timer for repeated holding
        else:
            vol_start_distance = None  # Reset if gesture released


    # Reset motion detection cooldown
    if motion_triggered:
        motion_timer -= 1
        if motion_timer <= 0:
            motion_triggered = False

    cv2.imshow("Gesture Control", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
