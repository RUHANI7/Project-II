# This app will use your built-in webcam to control your slides presentation.
# For a one-handed presentation, use Gesture 1 (thumbs up) to go to the previous slide and Gesture 2 (whole hand pointing up) to go to the next slide.

import win32com.client                                                      # For controlling PowerPoint via COM interface
from cvzone.HandTrackingModule import HandDetector                          # Hand tracking module
import cv2                                                                  # OpenCV for image processing
import numpy as np                                                          # NumPy for array operations

# PowerPoint setup
Application = win32com.client.Dispatch("PowerPoint.Application")
# Open the PowerPoint presentation file
Presentation = Application.Presentations.Open(
    "C:/Users/KIIT/Downloads/PPT-Presentation-controlled-by-hand-gesture-main/PPT-Presentation-controlled-by-hand-gesture-main/zani.pptx")
print(Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 900, 720                                                    # Width and height of the camera frame
gestureThreshold = 300                                                      # Threshold for hand position to detect gestures

# Camera Setup
cap = cv2.VideoCapture(0)                                                   # Initialize webcam capture
cap.set(3, width)                                                           # Set width of capture frame
cap.set(4, height)                                                          # Set height of capture frame

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)                   # Initialize hand detector

# Variables
delay = 30                                                                  # Delay counter to prevent rapid triggering
buttonPressed = False                                                       # Flag to indicate if a gesture has triggered an action
counter = 0                                                                 # Counter for delay
imgNumber = 20                                                              # Initial image number (not used in this code)
annotations = [[]]                                                          # List to store annotations (drawing paths)
annotationNumber = -1                                                       # Index for the current annotation being drawn
annotationStart = False                                                     # Flag to indicate if drawing annotation is active

while True:
    # Get image frame from camera
    success, img = cap.read()

    # Find hands in the frame and get hand landmarks
    hands, img = detectorHand.findHands(img)

    # If hands are detected and no gesture has been processed
    if hands and buttonPressed is False:
        hand = hands[0]                                                    # Consider only the first detected hand
        cx, cy = hand["center"]                                            # Center coordinates of the hand
        lmList = hand["lmList"]                                            # List of 21 landmark points of the hand
        fingers = detectorHand.fingersUp(hand)                             # Boolean list of which fingers are up

        # Check if hand is at the height of the face (gesture threshold)
        if cy <= gestureThreshold:
            # Gesture: Thumb up (Next Slide)
            if fingers == [1, 1, 1, 1, 1]:
                print("Next")
                buttonPressed = True
                Presentation.SlideShowWindow.View.Next()                   # Go to the next slide in PowerPoint
                annotations = [[]]                                         # Clear annotations
                annotationNumber = -1                                      # Reset annotation index
                annotationStart = False                                    # Disable annotation drawing mode

            # Gesture: One finger extended (Previous Slide)
            elif fingers == [1, 0, 0, 0, 0]:
                print("Previous")
                buttonPressed = True
                Presentation.SlideShowWindow.View.Previous()              # Go to the previous slide in PowerPoint
                annotations = [[]]                                        # Clear annotations
                annotationNumber = -1                                     # Reset annotation index
                annotationStart = False                                   # Disable annotation drawing mode

            # Gesture: Thumb and index finger extended (Zoom In)
            elif fingers == [0, 1, 0, 0, 0]:
                print("Zoom In")
                buttonPressed = True
                # Implement zoom in action (not implemented in this version)

            # Gesture: Thumb, index finger, and middle finger extended (Zoom Out)
            elif fingers == [0, 1, 1, 0, 0]:
                print("Zoom Out")
                buttonPressed = True
                # Implement zoom out action (not implemented in this version)

            # Gesture: Thumb, index finger, and ring finger extended (Draw Mode)
            elif fingers == [0, 1, 1, 1, 0]:
                print("Draw Mode")
                buttonPressed = True
                annotationStart = not annotationStart                     # Toggle drawing mode
                if annotationStart:
                    annotationNumber += 1                                 # Increment annotation index
                    annotations.append([])                                # Start a new annotation path

        # If in drawing annotation mode, draw on the image
        if annotationStart:
            if fingers == [0, 1, 1, 1, 0]:                                # Check if still in draw mode
                x, y = lmList[8][0], lmList[8][1]                         # Coordinates of the tip of index finger
                annotations[annotationNumber].append((x, y))              # Add point to current annotation path

    # Handle button pressed delay to prevent rapid triggers
    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    # Draw annotations on the image
    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(img, annotation[j - 1], annotation[j], (0, 0, 200), 12)  # Draw lines between points

    # Display the annotated image
    cv2.imshow("Image", img)

    # Check for 'q' key press to quit the program
    key = cv2.waitKey(1)
    if key == ord('q'):
        break

# Release the camera and close all OpenCV windows
cap.release()
cv2.destroyAllWindows()
