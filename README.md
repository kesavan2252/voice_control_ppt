# Voice-Controlled PowerPoint Presentation

This Python program allows you to control a PowerPoint presentation using voice commands. You can start the presentation, move to the next or previous slide, stop the presentation, and even close PowerPoint using voice recognition.

## Features

- Start the presentation with voice commands like "start", "launch", or "begin".
- Move to the next slide by saying "next", "forward", or "advance".
- Go to the previous slide by saying "previous", "back", or "go back".
- Stop the presentation with commands like "stop" or "end".
- Quit the program or close PowerPoint with voice commands.

## Requirements

- Python 3.x
- Internet connection (for Google Speech Recognition API)
- PowerPoint installed on your system

## Installation

1. **Clone the Repository**
   ```
   git clone https://github.com/your-username/voice-control-ppt.git
   cd voice-control-ppt
   ```

2. **Install Required Python Packages**
   You need to install several Python libraries before running the program. These include `pyttsx3` for text-to-speech, `speech_recognition` for voice recognition, and others.

   Use the following command to install all the required packages:

   ```
   pip install pyttsx3 SpeechRecognition pyautogui pygetwindow
   ```

   If you don't have `pip` installed, follow the [pip installation guide](https://pip.pypa.io/en/stable/installation/) first.

3. **Set the Path for Your PowerPoint File**
   In the script, replace the placeholder `ppt_path` with the actual path to your PowerPoint `.pptx` file:

   ```python
   ppt_path = r"C:\Users\YourUsername\Path\To\Your\Presentation.pptx"
   ```

4. **Run the Program**
   Once everything is set up, you can run the program:

   ```
   python voice_control_ppt.py
   ```

   The program will automatically open your PowerPoint file and start listening for voice commands.

## Usage

1. **Ensure Microphone Access**
   The program uses your microphone to recognize voice commands. Make sure that your microphone is working and accessible by Python.

2. **Give Commands**
   - **Start Presentation**: Say "start", "launch", or "begin".
   - **Next Slide**: Say "next", "forward", or "advance".
   - **Previous Slide**: Say "previous", "back", or "go back".
   - **Stop Presentation**: Say "stop" or "end".
   - **Quit Program**: Say "quit".
   - **Close PowerPoint**: Say "close".

3. **Note**: Make sure PowerPoint is the active window for the program to work properly. The script automatically brings PowerPoint to the front, but make sure nothing interferes with it.

## How It Works

- The program uses the `speech_recognition` library to capture voice input and convert it into text.
- Based on the recognized voice commands, it uses `pyautogui` to control PowerPoint (e.g., pressing the keys to start or move between slides).
- `pyttsx3` is used for text-to-speech, providing spoken feedback after each action.
- `pygetwindow` is used to ensure that PowerPoint is the active window when commands are issued.

## Troubleshooting

- If the program doesn't recognize your commands, make sure your microphone is working correctly and check your internet connection.
- If the program fails to focus on the PowerPoint window, make sure the PowerPoint application is open and running.
- You can modify the `total_slides` variable in the script to match the number of slides in your presentation for the best experience.
