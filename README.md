# ASL-System-Controller
Control The PPT Using American Sign Language

A real-time American Sign Language (ASL) gesture recognition system that controls PowerPoint presentations and system functions using hand gestures through your webcam.

🌟 Features

Real-time hand gesture recognition using MediaPipe and CVZone
PowerPoint presentation control (next/previous slides, jump, pause, exit)
System function control (volume, applications, screenshots, etc.)
Dual-hand detection for better accuracy
Two operating modes: Presentation and System Control
Gesture-based mode switching - no keyboard required
Camera auto-recovery if connection is lost
Cross-platform support (Windows focus, with some Linux/Mac compatibility)

🎥 Demo
The system recognizes 12+ different ASL gestures and can:

Control PowerPoint slides hands-free
Open applications and take screenshots
Switch between different control modes
Work with either left or right hand

🛠️ Installation
Prerequisites

Python 3.8+ (tested with Python 3.12.6)
Webcam (built-in or external)
Windows (for PowerPoint and some system functions)
PowerPoint (optional, for presentation mode)

!!The System Control mode is not fully functional!!

Install required packages:
bashpip install -r requirements.txt
requirements.txt
opencv-python==4.8.1.78
mediapipe==0.10.7
cvzone==1.6.1
numpy==1.26.2
pyautogui==0.9.54
pywin32==306
Pillow==10.1.0
protobuf==4.25.1
colorama==0.4.6
Setup Steps

Clone or download the project
bashgit clone <repository-url>
cd asl-system-controller

Create virtual environment
bashpython -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac

Install dependencies
bashpip install -r requirements.txt

Run the application
bashpython main.py


🚀 Usage
Starting the Application

With PowerPoint file:
python# Update the path in main.py
ppt_path = r"C:\path\to\your\presentation.pptx"
