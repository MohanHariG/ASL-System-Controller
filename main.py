import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import numpy as np
import pyautogui
import subprocess
import time

class SimpleASLController:
    def __init__(self):
        # PowerPoint Setup
        self.Application = None
        self.Presentation = None
        self.ppt_active = False

        # Camera Parameters
        self.width, self.height = 1200, 800
        
        # Initialize camera with better error handling
        self.cap = None
        self.init_camera()

        # Hand Detection - Allow 2 hands for better mode switching
        self.detectorHand = HandDetector(detectionCon=0.7, maxHands=2)

        # Control Variables
        self.buttonPressed = False
        self.counter = 0
        self.delay = 30
        self.paused = False
        self.mode = "presentation"  # modes: presentation, system
        
        # Gesture history for smoothing
        self.gesture_history = []
        self.history_size = 5

    def init_camera(self):
        """Initialize camera with multiple attempts and better error handling"""
        print("Initializing camera...")
        
        # Try multiple camera indices and backends
        camera_configs = [
            (0, cv2.CAP_DSHOW),  # DirectShow (Windows)
            (1, cv2.CAP_DSHOW),
            (0, cv2.CAP_ANY),    # Any available backend
            (1, cv2.CAP_ANY),
            (0, None),           # Default
            (1, None),
        ]
        
        for i, (index, backend) in enumerate(camera_configs):
            try:
                print(f"Trying camera {index} with backend {backend}...")
                
                if backend is not None:
                    cap = cv2.VideoCapture(index, backend)
                else:
                    cap = cv2.VideoCapture(index)
                
                if cap.isOpened():
                    # Test if we can actually read a frame
                    ret, frame = cap.read()
                    if ret and frame is not None:
                        self.cap = cap
                        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.width)
                        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.height)
                        self.cap.set(cv2.CAP_PROP_FPS, 30)
                        print(f"Successfully initialized camera {index}")
                        return
                    else:
                        cap.release()
                else:
                    cap.release()
                    
            except Exception as e:
                print(f"Failed to initialize camera {index}: {e}")
                continue
        
        raise Exception("No working camera found! Please check:\n"
                       "1. Camera is connected\n"
                       "2. Camera is not being used by another application\n"
                       "3. Camera drivers are installed")

    def reinit_camera(self):
        """Reinitialize camera if connection is lost"""
        if self.cap:
            self.cap.release()
        time.sleep(1)  # Wait a moment
        try:
            self.init_camera()
            return True
        except:
            return False

    def load_presentation(self, ppt_path):
        """Load and start PowerPoint presentation"""
        try:
            self.Application = win32com.client.Dispatch("PowerPoint.Application")
            self.Presentation = self.Application.Presentations.Open(ppt_path)
            self.Presentation.SlideShowSettings.Run()
            self.ppt_active = True
            print("PowerPoint presentation loaded successfully")
        except Exception as e:
            print(f"Error loading presentation: {e}")
            self.ppt_active = False

    def get_slide_info(self):
        """Get current and total slides"""
        if self.ppt_active and self.Presentation:
            try:
                current_slide = self.Presentation.SlideShowWindow.View.Slide.SlideIndex
                total_slides = self.Presentation.Slides.Count
                return current_slide, total_slides
            except:
                return 1, 1
        return 1, 1

    def recognize_hand_sign(self, hand):
        """Recognize hand signs using finger positions (works for both hands)"""
        fingers = self.detectorHand.fingersUp(hand)
        landmarks = hand["lmList"]
        
        # Get key landmark positions
        thumb_tip = landmarks[4]
        index_tip = landmarks[8]
        
        # Calculate distance for OK sign
        thumb_index_dist = np.sqrt((thumb_tip[0] - index_tip[0])**2 + (thumb_tip[1] - index_tip[1])**2)
        
        # Gesture recognition based on finger patterns (mirror-corrected)
        if fingers == [0, 0, 0, 0, 0]:
            return "FIST"
        elif fingers == [1, 0, 0, 0, 0]:
            return "THUMBS_UP"
        elif fingers == [0, 1, 0, 0, 0]:
            return "POINT"
        elif fingers == [0, 1, 1, 0, 0]:
            return "PEACE"
        elif fingers == [1, 1, 1, 1, 1]:
            return "FIVE"
        elif fingers == [0, 0, 0, 0, 1]:
            return "PINKY"
        elif fingers == [1, 0, 0, 0, 1]:
            return "ROCK_ON"
        elif fingers == [1, 1, 0, 0, 0]:
            return "TWO"
        elif fingers == [0, 1, 1, 1, 0]:
            return "THREE"
        elif fingers == [0, 1, 1, 1, 1]:
            return "FOUR"
        elif fingers == [1, 1, 0, 0, 1]:
            return "I_LOVE_YOU"
        elif fingers == [1, 0, 1, 0, 0]:
            return "MODE_SWITCH"
        elif fingers == [1, 1, 1, 0, 0]:
            return "THREE_FINGER_MODE"  # Alternative mode switch
        
        # OK sign based on distance
        if thumb_index_dist < 30 and fingers[0] == 1 and fingers[1] == 1:
            return "OK_SIGN"
        
        return "UNKNOWN"

    def smooth_gesture_recognition(self, gesture):
        """Smooth gesture recognition using history"""
        self.gesture_history.append(gesture)
        if len(self.gesture_history) > self.history_size:
            self.gesture_history.pop(0)
        
        # Return most common gesture in recent history
        if len(self.gesture_history) >= 3:
            most_common = max(set(self.gesture_history), key=self.gesture_history.count)
            return most_common
        return gesture

    def execute_presentation_action(self, label):
        """Execute presentation control actions"""
        if not self.ppt_active:
            print("No presentation loaded")
            return
            
        current_slide, total_slides = self.get_slide_info()
        
        if label == "FIVE":  # Next slide
            if current_slide < total_slides:
                self.Presentation.SlideShowWindow.View.Next()
                print(f"Next slide: {current_slide + 1}/{total_slides}")
            else:
                print("Already on last slide")
                
        elif label == "THUMBS_UP":  # Previous slide
            if current_slide > 1:
                self.Presentation.SlideShowWindow.View.Previous()
                print(f"Previous slide: {current_slide - 1}/{total_slides}")
            else:
                print("Already on first slide")
                
        elif label == "PEACE":  # Jump forward 2 slides
            jump_to = min(current_slide + 2, total_slides)
            self.Presentation.SlideShowWindow.View.GotoSlide(jump_to)
            print(f"Jumped to slide: {jump_to}/{total_slides}")
                
        elif label == "PINKY":  # Exit slideshow
            self.Presentation.SlideShowWindow.View.Exit()
            self.ppt_active = False
            print("Slideshow ended")
            
        elif label == "OK_SIGN":  # Go to first slide
            self.Presentation.SlideShowWindow.View.GotoSlide(1)
            print("Returned to first slide")

    def execute_system_action(self, label):
        """Execute system control actions"""
        if label == "POINT":  # Open calculator
            try:
                subprocess.Popen('calc.exe')
                print("Calculator opened")
            except:
                print("Could not open calculator")
                
        elif label == "OK_SIGN":  # Take screenshot
            try:
                screenshot = pyautogui.screenshot()
                timestamp = int(time.time())
                screenshot.save(f'screenshot_{timestamp}.png')
                print(f"Screenshot saved: screenshot_{timestamp}.png")
            except:
                print("Could not take screenshot")
                
        elif label == "ROCK_ON":  # Open task manager
            try:
                pyautogui.hotkey('ctrl', 'shift', 'esc')
                print("Task manager opened")
            except:
                print("Could not open task manager")
                
        elif label == "I_LOVE_YOU":  # Lock screen
            try:
                pyautogui.hotkey('win', 'l')
                print("Screen locked")
            except:
                print("Could not lock screen")
                
        elif label == "FIVE":  # Open notepad
            try:
                subprocess.Popen('notepad.exe')
                print("Notepad opened")
            except:
                print("Could not open notepad")

    def handle_mode_switching(self, label):
        """Handle mode switching with multiple options"""
        if label == "MODE_SWITCH" or label == "THREE_FINGER_MODE":
            modes = ["presentation", "system"]
            current_index = modes.index(self.mode)
            next_index = (current_index + 1) % len(modes)
            self.mode = modes[next_index]
            print(f"Switched to {self.mode.upper()} mode")
            return True
        return False

    def draw_ui(self, img, label):
        """Draw user interface elements"""
        # Mode indicator
        cv2.rectangle(img, (10, 10), (350, 60), (0, 0, 0), -1)
        cv2.putText(img, f"Mode: {self.mode.upper()}", (20, 35), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
        
        # Detected sign
        cv2.rectangle(img, (10, 70), (450, 120), (0, 0, 0), -1)
        cv2.putText(img, f"Sign: {label}", (20, 95), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 0), 2)
        
        # Slide info (if presentation mode)
        if self.mode == "presentation" and self.ppt_active:
            current, total = self.get_slide_info()
            cv2.rectangle(img, (10, 130), (300, 180), (0, 0, 0), -1)
            cv2.putText(img, f"Slide: {current}/{total}", (20, 155), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 0, 255), 2)
        
        # Camera status
        cv2.rectangle(img, (self.width - 200, 10), (self.width - 10, 60), (0, 100, 0), -1)
        cv2.putText(img, "Camera: Active", (self.width - 190, 35), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        
        # Instructions
        instructions = [
            "Controls:",
            "Thumb+Middle OR Three fingers: Switch Mode",
            "Use either hand - both detected",
            "Q: Quit, R: Restart camera"
        ]
        
        for i, text in enumerate(instructions):
            cv2.putText(img, text, (10, self.height - 80 + i*18), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)

    def run(self, ppt_path=None):
        """Main execution loop with camera recovery"""
        print("=== Simple ASL Controller Started ===")
        print("Controls:")
        print("- Show gestures with either hand")
        print("- Mode switching: Thumb+Middle finger OR Three fingers (thumb+index+middle)")
        print("- Both hands are detected for better recognition")
        print("- Press 'Q' to quit")
        print("- Press 'R' to restart camera if it fails")
        print("- Current mode:", self.mode.upper())
        
        if ppt_path:
            self.load_presentation(ppt_path)
        else:
            print("No PowerPoint file specified - system mode available")

        consecutive_failures = 0
        max_failures = 5

        while True:
            if not self.cap or not self.cap.isOpened():
                print("Camera connection lost. Attempting to reconnect...")
                if not self.reinit_camera():
                    print("Failed to reconnect camera. Press 'R' to retry.")
                    key = cv2.waitKey(1000) & 0xFF
                    if key == ord('q') or key == ord('Q'):
                        break
                    elif key == ord('r') or key == ord('R'):
                        continue
                    continue

            success, img = self.cap.read()
            
            if not success:
                consecutive_failures += 1
                print(f"Failed to read from camera (attempt {consecutive_failures})")
                
                if consecutive_failures >= max_failures:
                    print("Too many camera failures. Attempting to restart camera...")
                    if self.reinit_camera():
                        consecutive_failures = 0
                        continue
                    else:
                        print("Camera restart failed. Press 'R' to retry or 'Q' to quit.")
                        key = cv2.waitKey(1000) & 0xFF
                        if key == ord('q') or key == ord('Q'):
                            break
                        elif key == ord('r') or key == ord('R'):
                            consecutive_failures = 0
                            continue
                continue
            
            consecutive_failures = 0  # Reset failure counter on success

            # Don't flip image - this causes left/right confusion
            # img = cv2.flip(img, 1)
            
            # Detect hands (up to 2 hands now)
            hands, img = self.detectorHand.findHands(img)

            label = "NONE"
            
            if hands and not self.buttonPressed:
                # Process the most prominent hand or check all hands for mode switching
                processed = False
                
                # First check all hands for mode switching gestures
                for hand in hands:
                    gesture = self.recognize_hand_sign(hand)
                    if self.handle_mode_switching(gesture):
                        label = gesture
                        self.buttonPressed = True
                        processed = True
                        break
                
                # If no mode switching, process the first hand for other gestures
                if not processed and hands:
                    hand = hands[0]
                    raw_gesture = self.recognize_hand_sign(hand)
                    label = self.smooth_gesture_recognition(raw_gesture)
                    
                    if label != "UNKNOWN" and not self.handle_mode_switching(label):
                        # Execute action based on current mode
                        if self.mode == "presentation":
                            self.execute_presentation_action(label)
                        elif self.mode == "system":
                            self.execute_system_action(label)
                        
                        self.buttonPressed = True

            # Reset button press after delay
            if self.buttonPressed:
                self.counter += 1
                if self.counter > self.delay:
                    self.counter = 0
                    self.buttonPressed = False

            # Draw UI
            self.draw_ui(img, label)

            # Display image
            cv2.imshow("Simple ASL Controller", img)

            # Handle key presses
            key = cv2.waitKey(1) & 0xFF
            if key == ord('q') or key == ord('Q'):
                print("Exiting...")
                break
            elif key == ord('r') or key == ord('R'):
                print("Restarting camera...")
                self.reinit_camera()

        self.cleanup()

    def cleanup(self):
        """Clean up resources"""
        print("Cleaning up...")
        if self.cap and self.cap.isOpened():
            self.cap.release()
        cv2.destroyAllWindows()
        
        if self.ppt_active and self.Application:
            try:
                self.Application.Quit()
                print("PowerPoint closed")
            except:
                pass

# Main execution
if __name__ == "__main__":
    try:
        controller = SimpleASLController()
        
        # Update this path to your PowerPoint file or set to None for system mode only
        ppt_path = r"C:\Users\G mohanhari\Desktop\sign\zani.pptx"
        
        # Check if file exists
        import os
        if ppt_path and os.path.exists(ppt_path):
            controller.run(ppt_path)
        else:
            print("PowerPoint file not found, running in system mode only")
            controller.run()
            
    except Exception as e:
        print(f"Error starting controller: {e}")
        input("Press Enter to exit...")