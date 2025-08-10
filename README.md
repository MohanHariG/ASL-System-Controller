# Gesture-PPT-Controller using ASL
Control The PPT Using American Sign Language

A real-time American Sign Language (ASL) gesture recognition system that controls PowerPoint presentations and system functions using hand gestures through your webcam.

🌟 Features

Real-time hand gesture recognition using MediaPipe and CVZone
PowerPoint presentation control (next/previous slides, jump, pause, exit)
System function control (volume, applications, screenshots, etc.)
Dual-hand detection for better accuracy
Camera auto-recovery if the connection is lost
Cross-platform support (Windows focus, with some Linux/Mac compatibility)

🎥 Demo
The system recognizes 12+ different ASL gestures and can:

Control PowerPoint slides hands-free
Work with either the left or the right hand

🛠️ Installation
Prerequisites

Python 3.8+ (tested with Python 3.12.6)
Webcam (built-in or external)
Windows (for PowerPoint and some system functions)
PowerPoint (optional, for presentation mode)

Clone or download the project
bashgit clone <repository-url>
cd asl-system-controller

Create a virtual environment
bashpython -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac

Install dependencies
pip install -r requirements.txt

Run the application
bashpython main.py


🚀 Usage
Starting the Application

With a PowerPoint file:
python# Update the path in main.py
ppt_path = r"C:\path\to\your\presentation.pptx"
