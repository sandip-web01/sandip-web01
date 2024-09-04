from flask import Flask, request, render_template
import speech_recognition as sr
import time
import datetime
import webbrowser
import os
from openpyxl import Workbook

app = Flask(__name__)

# Initialize the recognizer
recognizer = sr.Recognizer()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run-python', methods=['POST'])
def run_python():
    wish_me()
    run_once = 1
    query = ""
    while run_once == 1:
        run_once += 1
        query = listen()
        if query in ["exit", "stop"]:
            speak("Thank you. You are a good speaker. Goodbye! Have a nice time.")
            break
        perform_task(query)
    return render_template('index.html', command=query)

def listen(slowdown_factor=1):
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

        try:
            time.sleep(slowdown_factor)
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language='en-US')
            print(f"User said: {query}")
            return query.lower()
        
        except sr.UnknownValueError:
            speak("Sorry, I did not catch that. Could you please repeat?")
        except sr.RequestError:
            speak("Sorry, the service is down. Please try again later.")
        return ""

def wish_me():
    hour = datetime.datetime.now().hour
    greeting = "Good Morning!" if hour < 12 else "Good Afternoon!" if hour < 18 else "Good Evening!"
    speak(greeting)
    speak("I am your personal assistant. How can I help you today?")

def perform_task(query):
    if "time" in query:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"The time is {current_time}")
    elif "search for" in query:
        search_term = query.replace("search for", "").strip()
        speak(f"Searching for {search_term} on Google")
        webbrowser.open(f"https://www.google.com/search?q={search_term.replace(' ', '+')}")
    elif "play" in query and "youtube" in query:
        song_name = query.replace("play", "").replace("on youtube", "").strip()
        speak(f"Playing {song_name} on YouTube")
        webbrowser.open(f"https://www.youtube.com/results?search_query={song_name.replace(' ', '+')}")
    elif "open" in query:
        website = query.replace("open", "").strip().replace(" ", "")
        speak(f"Opening {website}")
        webbrowser.open(f"https://{website}.com")
    elif "say" in query:
        speak(query.replace("say", "").strip())
    else:
        speak("Sorry, I can not help with that right now.")

def speak(text):
    # Simple print statement for Vercel compatibility
    print(f"Speaking: {text}")

if __name__ == "__main__":
    app.run()
