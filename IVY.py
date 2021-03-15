from __future__ import print_function
import datetime
import pickle
import requests	 
from output_module import output
from internet import check_internet_connection
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechrecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser#pip install webbrowser
import os#pip install os
#pip install pipwin
#pipwin install pyaudio
engine = pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
#print(voices[1].id)
engine.setProperty("voice",voices[1].id)
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
DAY_EXTENSIONS = ["nd", "rd", "th", "st"]
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Hi..Good Morning!")
    elif hour>=12 and hour<=16:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("Hi..........................................Iam IVY.Tell me how can I help you  ")
def sendEmail(to, content): 
    server = smtplib.SMTP('smtp.gmail.com', 587) 
    server.ehlo() 
    server.starttls() 
      
    # Enable low security in gmail 
    server.login('ivybot26@gmail.com', 'MKSTL@123') 
    server.sendmail('ivybot26@gmail.com', to, content) 
    server.close() 
def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening......")
        speak("Listening......")

        r.pause_threshold = 0.8
        audio=r.listen(source)

    try:
        print("Recognizing........")
        query = r.recognize_google(audio, language="en-in")
        print(f"{query}\n")
    except Exception as e:
        # print(e)
        speak("I didn't get you.....")
        speak("Can u please Repeat...?")
        print("I didn't get you.....")
        return "None"
    return query
if __name__ == "__main__":
    clear = lambda:os.system('cls')
    clear()
    wishMe()
    takeCommand()
    while True:
        query = takeCommand().lower()
        if 'wikipedia' in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query,sentences=2)
            speak("According to wikipedia")
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        elif 'open facebook' in query:
            webbrowser.open("facebook.com")
        elif 'your name' in query:
            speak("Iam IVY.....Iam here to help you!")
        elif 'open Instagram' in query:
            webbrowser.open("instagram.com")
        elif 'open google' in query:
            webbrowser.open("google.com")
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")
            print(strTime)
        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("Whom should I send?")
                to = input()
                sendEmail(to,content)
                speak("I have delivered the Mail")
            except Exception as e:
                print(e)
                speak("I couldn't Deliver the Mail")
        elif 'shutdown system' in query: 
                speak("I am shutting down the system") 
                subprocess.call('shutdown / p /f')
        elif "restart" in query: 
            subprocess.call(["shutdown", "/r"])
        elif 'lock window' in query: 
                speak("locking the device") 
                ctypes.windll.user32.LockWorkStation()
        elif 'exit' in query: 
            speak("Bye") 
            exit()
        elif "how are you" in query: 
            speak("I'm fine, glad you asked me that") 
        elif "i love you" in query: 
            speak("Why don't we change the subject") 
def authenticate_google():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

def get_events(n, service):
    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print(f'Getting the upcoming {n} events')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])

def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    day = -1
    day_of_the_week = -1
    month = -1
    year= today.year
    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(month) + 1
        elif word in DAYS:
            day_of_the_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENSIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
service = authenticate_google()
get_events(2, service)

def get_news(): 
    if check_internet_connection():
        # BBC news api 
        main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=06d738298c994e01bd24f76b2caa8038"

        # fetching data in json format 
        open_bbc_page = requests.get(main_url).json() 

        # getting all articles in a string article 
        article = open_bbc_page["articles"] 

        # empty list which will 
        # contain all trending news 
        results = [] 
        
        for ar in article: 
            results.append(ar["title"]) 
            
        for i in range(len(results)): 
            
            # printing all trending news 
            output(str(i + 1)+ " ", results[i]) 

        return "So these were the top news from today"
    else:
         return "Please check your internet connection"