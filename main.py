# This program will only work on windows if pywin32 module is installed

import win32com.client as wincom
import time

speak = wincom.Dispatch("SAPI.SpVoice")

names = ["Arnav", "Coding", "Programming", "Python", "PyCharm", "JetBrains"]

for name in names:
    speak.Speak(f"Hello {name}")
    print(f"Said 'Hello {name}' ")
    # time.sleep(0.5) # uncomment this if you want some gap in between the speech

speak.Speak("Said hello to everyone!")
print("Given Shoutout to everyone!")
speak.Speak("Now Bye!")
print("Bye!")
