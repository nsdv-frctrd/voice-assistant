import subprocess
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
import os
import spotipy
import pygame
from spotipy.oauth2 import SpotifyOAuth
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
        engine.say(audio)
        engine.runAndWait()

def sing():
        sp = spotipy.Spotify(auth_manager=SpotifyOAuth(client_id='1f044e7b8bbf421d8eb238e8e279d00e',
                                               client_secret='2e27474b3c46410b911784c50fc7ead2',
                                               redirect_uri='https://github.com/nsdv-frctrd/voice-assistant',
                                               scope='playlist-read-private'))
        playlist_id = 'https://open.spotify.com/playlist/3fQ6EJdy6n1kF4Yw5bTAVx?si=cea0a37833014a67'
        playlist_tracks = sp.playlist_tracks(playlist_id)

        track_uris = [track['track']['uri'] for track in playlist_tracks['items']]
        pygame.init()
        pygame.mixer.init()

        for track_uri in track_uris:
            track_info = sp.track(track_uri)
            print(f"Now playing: {track_info['name']} by {track_info['artists'][0]['name']}")
            pygame.mixer.music.play()


def wishMe():
        speak("Greetings User")
        assistant_name =("Voice Assistant by Khooshi")
        speak("I am your Assistant")
        speak(assistant_name)
        

def username():
        speak("How would you like me to address you")
        uname = takeCommand()
        speak("Welcome")
        speak(uname)
        columns = shutil.get_terminal_size().columns
        print("Welcome ", uname.center(columns))        
        speak("How may I help you")

def takeCommand():
        
        r = sr.Recognizer()
        
        with sr.Microphone() as source:
                
                print("Listening...")
                r.pause_threshold = 1
                audio = r.listen(source)

        try:
                print("Recognizing...")
                query = r.recognize_google(audio, language ='en-in')
                print(f"User said: {query}\n")

        except Exception as e:
                print(e)
                print("Unable to recognize speech")
                return "None"
        
        return query

def sendEmail(to, content):
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        
        # Enable low security in gmail
        server.login('your email id', 'your email password')
        server.sendmail('your email id', to, content)
        server.close()

if __name__ == '__main__':
        clear = lambda: os.system('cls')
        wishMe()
        username()
        
        while True:
                
                query = takeCommand().lower()
                if 'open youtube' in query:
                        speak("Here you go to Youtube\n")
                        webbrowser.open("youtube.com")

                elif 'open google' in query:
                        speak("Here you go to Google\n")
                        webbrowser.open("google.com")

                elif 'open stackoverflow' in query:
                        speak("Here you go to Stack Overflow")
                        webbrowser.open("stackoverflow.com")

                elif 'play music' in query or "play song" in query:
                        speak("Here you go with music")
                        sing()

                elif 'the time' in query:
                        strTime = datetime.datetime.now().strftime("% H:% M:% S")
                        speak(f"Sir, the time is {strTime}")

                elif 'send a mail' in query:
                        try:
                                speak("What should I say?")
                                content = takeCommand()
                                speak("whome should i send")
                                to = input()
                                sendEmail(to, content)
                                speak("Email has been sent !")
                        except Exception as e:
                                print(e)
                                speak("I am not able to send this email")

                elif 'how are you' in query:
                        speak("I am fine, Thank you")
                        speak("How are you", )

                elif 'fine' in query or "good" in query:
                        speak("It's good to know that your fine")

                elif "change my name to" in query:
                        query = query.replace("change my name to", "")
                        assistant_name = query

                elif "change name" in query:
                        speak("What would you like to call me, Sir ")
                        assistant_name = takeCommand()
                        speak("Thanks for naming me")

                elif "what's your name" in query or "What is your name" in query:
                        speak("My friends call me")
                        speak(assistant_name)
                        print("My friends call me", assistant_name)

                elif 'exit' in query:
                        speak("Thanks for giving me your time")
                        exit()

                elif "who made you" in query or "who created you" in query:
                        speak("I have been created by Khooshi")
                        
                elif 'joke' in query:
                        speak(pyjokes.get_joke())
                        
                elif "calculate" in query:
                        
                        app_id = "Wolframalpha api id"
                        client = wolframalpha.Client(app_id)
                        indx = query.lower().split().index('calculate')
                        query = query.split()[indx + 1:]
                        res = client.query(' '.join(query))
                        answer = next(res.results).text
                        print("The answer is " + answer)
                        speak("The answer is " + answer)

                elif 'search' in query or 'play' in query:
                        
                        query = query.replace("search", "")
                        query = query.replace("play", "")               
                        webbrowser.open(query)

                elif "who i am" in query:
                        speak("If you talk then definitely your human.")

                elif "why you came to world" in query:
                        speak("Thanks to Gaurav. further It's a secret")

                elif "who are you" in query:
                        speak("I am your virtual assistant created by Khooshi")

                elif 'change background' in query:
                        ctypes.windll.user32.SystemParametersInfoW(20,
                                                                                                        0,
                                                                                                        "Location of wallpaper",
                                                                                                        0)
                        speak("Background changed successfully")

                elif 'news' in query:
                        
                        try:
                                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                                data = json.load(jsonObj)
                                i = 1
                                
                                speak('here are some top news from the times of india')
                                print('''=============== TIMES OF INDIA ============'''+ '\n')
                                
                                for item in data['articles']:
                                        
                                        print(str(i) + '. ' + item['title'] + '\n')
                                        print(item['description'] + '\n')
                                        speak(str(i) + '. ' + item['title'] + '\n')
                                        i += 1
                        except Exception as e:
                                
                                print(str(e))

                
                elif 'lock window' in query:
                                speak("locking the device")
                                ctypes.windll.user32.LockWorkStation()

                elif 'shutdown system' in query:
                                speak("Hold On a Sec ! Your system is on its way to shut down")
                                subprocess.call('shutdown / p /f')
                                
                elif 'empty recycle bin' in query:
                        winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                        speak("Recycle Bin Recycled")

                elif "don't listen" in query or "stop listening" in query:
                        speak("for how much time you want to stop jarvis from listening commands")
                        a = int(takeCommand())
                        time.sleep(a)
                        print(a)

                elif "where is" in query:
                        query = query.replace("where is", "")
                        location = query
                        speak("User asked to Locate")
                        speak(location)
                        webbrowser.open("https://www.google.nl / maps / place/" + location + "")

                elif "camera" in query or "take a photo" in query:
                        ec.capture(0, "Jarvis Camera ", "img.jpg")

                elif "restart" in query:
                        subprocess.call(["shutdown", "/r"])
                        
                elif "hibernate" in query or "sleep" in query:
                        speak("Hibernating")
                        subprocess.call("shutdown / h")

                elif "log off" in query or "sign out" in query:
                        speak("Make sure all the application are closed before sign-out")
                        time.sleep(5)
                        subprocess.call(["shutdown", "/l"])

                elif "write a note" in query:
                        speak("What should i write, sir")
                        note = takeCommand()
                        file = open('jarvis.txt', 'w')
                        strTime = datetime.datetime.now().strftime("% H:% M:% S")
                        file.write(strTime)
                        file.write(" :- ")
                        file.write(note)
                
                elif "show note" in query:
                        speak("Showing Notes")
                        file = open("assistant.txt", "r")
                        print(file.read())
                        speak(file.read(6))

                elif "assistant" in query:
                        
                        wishMe()
                        speak("I am Voice Assistant by Khooshi")
                        speak(assistant_name)

                elif "weather" in query:
                        api_key = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
                        base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
                        speak(" City name ")
                        print("City name : ")
                        city_name = takeCommand()
                        complete_url = base_url + "appid =" + api_key + "&q =" + city_name
                        response = requests.get(complete_url)
                        x = response.json()
                        
                        if x["code"] != "404":
                                y = x["main"]
                                current_temperature = y["temp"]
                                current_pressure = y["pressure"]
                                current_humidiy = y["humidity"]
                                z = x["weather"]
                                weather_description = z[0]["description"]
                                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                        
                        else:
                                speak(" City Not Found ")
                        
                
                elif "wikipedia" in query:
                        webbrowser.open("wikipedia.com")

                elif "Good Morning" in query:
                        speak("A warm" +query)
                        speak("How are you ")
                        speak(assistant_name)

                # most asked question from google Assistant
                
                elif "how are you" in query:
                        speak("I'm fine, glad you me that")

                elif "what is" in query or "who is" in query:
                        client = wolframalpha.Client("NPPR9-FWDCX-D2C8J-H872K-2YT43")
                        res = client.query(query)
                        
                        try:
                                print (next(res.results).text)
                                speak (next(res.results).text)
                        except StopIteration:
                                print ("No results")


