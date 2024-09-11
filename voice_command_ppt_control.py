import pyttsx3
import speech_recognition as sr
import os
import pyautogui
import pygetwindow as gw
import time

# Initialize speech engine
engine = pyttsx3.init()

def speak(audio):
    """Speak out the given text."""
    engine.say(audio)
    engine.runAndWait()

def takeCommand():
    """Listen for a single word hotword using Google Speech Recognition."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for command...")
        r.adjust_for_ambient_noise(source)  # Adjust for background noise
        try:
            audio = r.listen(source)
            query = r.recognize_google(audio, language='en-in')
            print(f"Recognized Query: {query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Could not understand the audio")
            speak("Sorry, I did not get that. Could you please repeat?")
            return "None"
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            speak("Speech recognition service is unavailable. Please check your internet connection.")
            return "None"
        except Exception as e:
            print(f"Error recognizing the voice: {e}")
            speak("Say that again please...")
            return "None"

def controlPpt(total_slides):
    """Control PowerPoint with hotwords and detect last slide."""
    speak("Controlling PPT Presentation")
    focus_ppt_window()  # Bring PowerPoint to the front

    slide_number = 1  # Start from the first slide

    while True:
        query1 = takeCommand()

        # Handle multiple start commands
        if query1 in ["start", "launch", "begin"]:
            pyautogui.press('f5')  # Start presentation
            speak("Presentation started")
            time.sleep(5)  # Wait for 5 seconds before proceeding

        elif query1 in ["next", "forward", "advance"]:
            if slide_number < total_slides:
                pyautogui.press('space')  # Next slide
                slide_number += 1
                speak(f"Moving to slide {slide_number}")
                time.sleep(5)  # Wait for 5 seconds to read the slide
            else:
                speak("This is the last slide. Can't move forward.")

        elif query1 in ["previous", "back", "go back"]:
            if slide_number > 1:
                pyautogui.press('left')  # Previous slide
                slide_number -= 1
                speak(f"Going back to slide {slide_number}")
                time.sleep(5)  # Wait for 5 seconds to read the slide
            else:
                speak("This is the first slide.")

        elif query1 in ["stop", "end"]:
            pyautogui.press('esc')  # Stop presentation
            speak("Presentation stopped")
            break

        elif query1 == "quit":
            speak("Quitting the program")
            quit()

        elif query1 == "close":
            speak("Closing PowerPoint")
            os.system("TASKKILL /F /IM POWERPNT.exe")  # Close PowerPoint
            break

def focus_ppt_window():
    """Bring PowerPoint window to the front."""
    try:
        ppt_window = gw.getWindowsWithTitle('PowerPoint')[0]
        ppt_window.activate()
        print("PowerPoint window focused.")
        time.sleep(2)  # Ensure the window is focused
    except Exception as e:
        print(f"Error: {e}. Make sure PowerPoint is open.")

def open_ppt(ppt_path):
    """Open PowerPoint file automatically."""
    os.startfile(ppt_path)  # Opens the PowerPoint file
    time.sleep(10)  # Wait for PowerPoint to fully open

if __name__ == "__main__":
    total_slides = 10  # Replace this with the actual total number of slides
    ppt_path = r"C:\Users\Tamilarasu\Downloads\Voice_command_ppt.pptx"  # Replace with your PPT path
    open_ppt(ppt_path)  # Open PowerPoint automatically
    controlPpt(total_slides)
