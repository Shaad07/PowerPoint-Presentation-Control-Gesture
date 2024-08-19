import ctypes
import win32com.client
import win32gui
import win32con
import time
import mediapipe as mp
import cv2
import math

# Minimize the command prompt window
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32
hWnd = kernel32.GetConsoleWindow()
if hWnd:
    user32.ShowWindow(hWnd, 6)  # 6 = Minimize window

# Initialize PowerPoint application
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
presentation = powerpoint.Presentations.Open(r'C:\Users\shaad\Downloads\MP.pptx')  # Update the path
presentation.SlideShowSettings.Run()

# Bring PowerPoint window to the foreground
def find_powerpoint_window():
    time.sleep(2)
    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            windows.append(hwnd)
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    for hwnd in windows:
        if win32gui.GetWindowText(hwnd):
            if 'PowerPoint' in win32gui.GetWindowText(hwnd):
                return hwnd
    return None

hwnd = find_powerpoint_window()
if hwnd:
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
else:
    print("Error: PowerPoint window not found.")
    exit(1)

# Initialize MediaPipe for hand tracking
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(min_detection_confidence=0.7, min_tracking_confidence=0.5)
mp_drawing = mp.solutions.drawing_utils

# Open a webcam feed
cap = cv2.VideoCapture(0)

if not cap.isOpened():
    print("Error: Could not open webcam.")
    exit()

# Flag to track whether to close PowerPoint
close_presentation = False

while True:
    ret, frame = cap.read()
    if not ret:
        print("Error: Failed to capture image.")
        break

    frame = cv2.flip(frame, 1)
    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb_frame)

    hand_detected = False
    pinch_detected = False

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

            # Pinch Gesture Detection
            thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
            index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]

            # Calculate the Euclidean distance between the thumb tip and index finger tip
            thumb_index_distance = math.sqrt(
                (thumb_tip.x - index_tip.x) ** 2 +
                (thumb_tip.y - index_tip.y) ** 2 +
                (thumb_tip.z - index_tip.z) ** 2
            )

            # Debug: Print distance to help adjust the threshold
            print(f"Thumb-Index Distance: {thumb_index_distance}")

            # Define a threshold for detecting the pinch gesture
            pinch_threshold = 0.03  # You may need to adjust this value

            if thumb_index_distance < pinch_threshold:
                print("Pinch Gesture Detected")
                pinch_detected = True

            # Swipe Right Gesture: Thumb is to the right of the index finger
            thumb_tip_x = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP].x
            index_tip_x = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP].x
            if thumb_tip_x > index_tip_x:
                print("Swipe Right Gesture Detected: Next Slide")
                powerpoint.SlideShowWindows(1).View.Next()
                time.sleep(1)
            # Swipe Left Gesture: Index finger is to the right of the thumb
            elif index_tip_x > thumb_tip_x:
                print("Swipe Left Gesture Detected: Previous Slide")
                powerpoint.SlideShowWindows(1).View.Previous()
                time.sleep(1)

            hand_detected = True

    # if not hand_detected:
    #     print("No hands detected.")

    # Display the frame
    cv2.imshow('Hand Gesture Control', frame)

    # Check if the 'Esc' key is pressed
    if cv2.waitKey(1) & 0xFF == 27:
        break

    # Close presentation if the pinch_detected flag is set
    if pinch_detected:
        print("Closing PowerPoint presentation.")
        cap.release()
        cv2.destroyAllWindows()
        presentation.Close()
        powerpoint.Quit()
        hands.close()
        exit(0)

# Release resources if loop is exited manually
cap.release()
cv2.destroyAllWindows()
presentation.Close()
powerpoint.Quit()
hands.close()
