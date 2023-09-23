import json
import requests

def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    speak('Welcome To Listen News')
    speak('Which Category News You want to listen')
    print("Which Category News You want to listen")
    print("1. Top Headlines Of INDIA\n2. Business\n3. Entertainment\n4. Health\n"
          "5. Science\n6. Sports\n7. Technology ")
    speak('1 Top Headlines Of INDIA''2 Business''3 Entertainment''4 Health''5 Science''6 Sports''7 Technology')
    input_category = input("Choose Category ----> ")

    if input_category == '1':
        print('Top Headlines Of INDIA Are:')
        speak('Top Headlines Of INDIA Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")

    elif input_category == '2':
        print('Business News Are:')
        speak('Business News Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&category=business&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")

    elif input_category == '3':
        print('Entertainment News Are:')
        speak('Entertainment News Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&category=entertainment&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")

    elif input_category == '4':
        print('Health News Are:')
        speak('Health News Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&category=health&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")

    elif input_category == '5':
        print('Science News Are:')
        speak('Science News Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&category=science&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")

    elif input_category == '6':
        print('Sports News Are:')
        speak('Sports News Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&category=sports&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")

    elif input_category == '7':
        print('Technology News Are:')
        speak('Technology News Are')
        url = "https://newsapi.org/v2/top-headlines?country" \
              "=in&category=technology&apiKey=82dc81aa96d9428faac960891f39f887"
        news = requests.get(url).text
        news_json = json.loads(news)
        print(news_json["articles"])
        news_articles = news_json["articles"]
        for article in news_articles:
            print('TITLE')
            speak('Title')
            print(article['title'])
            speak(article['title'])
            print('DESCRIPTION')
            speak('Description')
            print(article['description'])
            speak(article['description'])
            print('\n')
            speak("Moving on next News........")
        speak("Thanks For listening")
