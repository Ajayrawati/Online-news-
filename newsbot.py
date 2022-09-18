import json
import requests


def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SPVoice")
    speak.Speak(str)

if __name__ == '__main__':
    data = requests.get('https://newsapi.org/v2/top-headlines?country=us&apiKey=dba7cd19c15f40e8bc258fd941814688')
    result = data.json()
    #print(result)
    
    news = result['articles']
    # speak(news['description'])
    speak("Welcome to news world")
    speak("So our first news is")

    for i in range(0, 10):
        print(news[i]['description'])
        speak(news[i]['description'])
        if i == 8:
            speak("So moving to our last news")
            continue
        if i == 9:
            break

    speak("Thanks")
