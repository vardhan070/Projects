import pyttsx3 
import speech_recognition as sr #saves time by speaking instead of typing
import datetime
import wikipedia #library to access data in wikipedia
import webbrowser #provides high level document interface to allow displaying web based documents to users
import os #provides functions to intract with OS
import smtplib #To send mails to any internet device
import random
import pyjokes
import win32com.client as wincl 
import ctypes #provides C compatable data types
import urlopen
import time
import subprocess

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait() #make the speech audible


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("hey vardhan! jarvis here, Please tell me how may I help you")       

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.8
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        speak("parden?")
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("opening youtube....")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("opening google....")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("opening stackoverflow....")
            webbrowser.open("stackoverflow.com")  

        elif 'open instagram' in query:
            speak("opening instagram....")
            webbrowser.open("instagram.com")     


        elif 'play music' in query or 'play songs' in query:
            music_dir = 'D:\\Downloads\\Downloads\\Music'
            songs = os.listdir(music_dir)
            z=random.choice(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[9]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open pycharm' in query:
            codePath = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2020.1.2\\bin\\pycharm64.exe"
            os.startfile(codePath)

        elif 'email to vardhan' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "iamvd7@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir, I am not able to send this email")   

        elif "hello jarvis" in query or "hi jarvis" in query:
            speak("Hello sir! How may I help you")        

        elif 'how are you' in query: 
            speak("I am fine, Thank you") 
            speak("How are you, Sir")

        elif "fine" in query: 
            speak("It's good to know that your fine") 

        elif  "not good" in query: 
            speak("Everything will be alright! dont worry just smile...")    

        elif "what's your name" in query or "what is your name" in query: 
            speak("my name is Jarvis,thats what my creator calls me!")  

        elif "who made you" in query or "who created you" in query or "who is your creator" in query:  
            speak("originally i was created by iron man in a movie, but here I have been created by Vardhan. He copied the name...! dont tell anyone, let it be a secret between us ")     

        elif "joke" in query:
            speak(pyjokes.get_joke())    

        elif "who am I" in query or "do you know me" in query:
            speak("Now you are talking to me then definately your vardhan or somehow related to vardhan")    
  
        elif "will you be my boyfriend" in query or "will you be my bf" in query:    
            speak("I'm not sure about. If u want to have one then u can ask the person who created me. He is single")  
  
        elif "i love you" in query: 
            speak("It's hard to understand. Ask my creator he will answer!")     
        
        elif "lock window" in query or "lock device" in query or "lock computer" in query:
                speak("locking the device") 
                ctypes.windll.user32.LockWorkStation()

        elif "write a note" in query or "take a note" in query: 
            speak("What should i write, sir") 
            note = takeCommand() 
            file = open('jarvis.txt', 'w') 
            file.write(note) 
            speak("done..!,you can check a new text folder has been created")
          
        elif "show note" in query: 
            speak("Showing Notes") 
            file = open("jarvis.txt", "r")  
            print(file.read()) 
            speak(file.read(6))  

        elif "don't listen" in query or "stop listening" in query: 
            speak("for how much time you want to stop jarvis from listening commands") 
            a = int(takeCommand()) 
            time.sleep(a) 
            print(a)    

        elif "open maps" in query:  
            speak("opening google maps") 
            webbrowser.open("https://www.google.co.in/maps/@16.9149533,56.9514475,3z")     

        elif "calculator"in query or "calculate" in query:
            speak("give me the values sir")
            pass
         
                
        if 'quit' in query or 'Bye' in query:
            speak("ok! Assistant jarvis signing-off. If you need anything then u can call me again")
            exit()       

