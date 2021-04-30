import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import json
import webbrowser
import subprocess
import wolframalpha
import tkinter
import operator
import winshell
import os
import sys
import spotipy
import spotipy.util as util
import smtplib
import random
import feedparser
import ctypes
import pyjokes
import time
import requests
import shutil
from json.decoder import JSONDecodeError
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am Jarvis. How may I help you sir")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold = 400
        r.dynamic_energy_threshold = True
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print("Say that again please...")
        return "None"
    return query


'''def mykspotify():
    username = sys.argv[1]
    scope = 'user-read-private user-read-playback-state user-modify-playback-state'
    try:
        token = util.prompt_for_user_token(username, scope)
    except (AttributeError, JSONDecodeError):
        os.remove(f".cache-{username}")
        token = util.prompt_for_user_token(username, scope)

    spotifyObject = spotipy.Spotify(auth=token)
    devices = spotifyObject.devices()
    print(json.dumps(devices, sort_keys=True, indent=4))
    deviceID = devices['devices'][0]['id']'''


def sendEmail(to, content):
    e = pass1.getEmail()
    p = pass1.getPass()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(e, p)
    server.sendmail(e, to, content)
    server.close()


if __name__ == "__main__":
    wishMe()
    while True:
        # if 1:
        query = takeCommand().lower()
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace('wikipedia', '')
            results = wikipedia.summary(query, sentences=3)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'Good Morning' in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak("Mayank")

        elif 'weather' in query:
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " + str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(
                    current_pressure) + "\n humidity (in percentage) = " + str(current_humidiy) + "\n description = " + str(weather_description))
            else:
                speak(" City Not Found ")

        elif 'who are you' in query:
            speak('I am jarvis sir')

        elif 'open youtube' in query:
            speak('opening youtube...')
            webbrowser.open("youtube.com")

        elif 'according to google' in query:
            speak('opening google...')
            query = query.replace('according to google', '')
            webbrowser.open("http://google.com/#q="+query, new=2)

        elif 'open github' in query:
            speak('opening github...')
            webbrowser.open("github.com")

        elif 'open netflix' in query:
            speak('opening netflix...')
            webbrowser.open("netflix.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif 'open amazon prime video' in query or 'open prime video' in query or 'open prime' in query:
            speak('opening amazon prime video...')
            webbrowser.open("primevideo.com")

        elif 'joke' in query:
            jok = pyjokes.get_joke()
            print(jok)
            speak(jok)

        elif 'play music' in query or 'play some music' in query:
            music_dir = "F:\\Music"
            songs = os.listdir(music_dir)
            randomSong = random.choice(songs)
            os.startfile(os.path.join(music_dir, randomSong))

        elif 'who made you' in query or 'who created you' in query:
            speak("I have been created by Mayank Bangwal.")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'calculate' in query:
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'open code' in query:
            speak('opening Visual Code Studio...')
            codePath = "C:\\Users\\rockm\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'news' in query:
            try:
                jsonObj = urlopen(
                    '''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============''' + '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'search' in query or 'play' in query:
            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.open(query)

        elif 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "rockmb972@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir, I am not able to send the email right now!")

        elif 'open valorant' in query or 'play valorant' in query:
            valorant = "C:\\Riot Games\\Riot Client\\RiotClientServices.exe"
            os.startfile(valorant)

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'where is' in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open(
                "https://www.google.nl / maps / place/" + location + "")

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif 'camera' in query or 'take a photo' in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif 'restart' in query:
            subprocess.call(["shutdown", "/r"])

        elif 'hibernate' in query or 'sleep' in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif 'write a note' in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif 'show note' in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read())

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif 'i love you' in query:
            speak("It's hard to understand")

        elif 'how are you' in query:
            speak("I'm fine sir, How are you?")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mayank")

        elif 'why you came to the world' in query:
            speak("Thanks to Mayank. Further It's a secret")

        elif 'dont listen' in query or 'stop listening' in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif 'bye' in query or 'goodbye' in query or 'turn off' in query:
            speak('goodbye sir')
            exit()
