import datetime
import os
import speech_recognition as sr
import win32com.client
import webbrowser
from config import apikey

open.api_key = apikey

def say(text):
    speaker = win32com.client.Dispatch("Sapi.SpVoice")
    # print("Enter Text You Want to Hear")
    # s = input()
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred Sorry"


if __name__ == '__main__':
    print('PyCharm')
    print("Hello I am Veda AI")
    say("Hello I am Veda AI")
    while True:
        print("Listening...")
        userInput = takeCommand()
        websites = [
            ["Google", "https://www.google.com"],
            ["YouTube", "https://www.youtube.com"],
            ["Facebook", "https://www.facebook.com"],
            ["Twitter", "https://www.twitter.com"],
            ["Instagram", "https://www.instagram.com"],
            ["Amazon", "https://www.amazon.com"],
            ["LinkedIn", "https://www.linkedin.com"],
            ["GitHub", "https://www.github.com"],
            ["Reddit", "https://www.reddit.com"],
            ["Wikipedia", "https://www.wikipedia.org"],
        ]
        for website in websites:
            if f"open {website[0]}".lower() in userInput.lower():
                say(f"Opening {website[0]} ")
                webbrowser.open(website[1])

        """app_list = [
            ["word", r"\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"],
            ["excel", r"\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk"],
            ["powerpoint", r"\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk"],
        ]

        for app in app_list:
            if f"open {app[0]} application" in userInput:
                say("Opening {app}")
                os.system(f"open {app[1]}")"""

        if "Play music" in userInput:
            #say("opening music sir")
            musicpath = "\\Users\\athar\\Music\\titanium-170190.mp3"
            #os.system(f"open {musicpath}")
            #os.system(musicpath)
            webbrowser.open(musicpath)
        #say(userInput)

        if "the time" in userInput:
            time = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"The Time is {time}")
            say(f"The Time is {time}")
