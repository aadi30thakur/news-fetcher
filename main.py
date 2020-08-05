from win32com.client import Dispatch
import requests
country = input(
    "Please enter the country name of which you wanted to see news : \n")
url = "http://newsapi.org/v2/top-headlines?" + "country=" + \
    country + '&' + "apiKey=45b3068e851842fbbf3fe28772d995b9"
webpage_asJson = requests.get(url).json()
articles = webpage_asJson["articles"]


results = []

for art in articles:
    results.append(art['title'])

for i in range(len(results)):
    print(i+1, results[i])


speak = Dispatch("SAPI.SpVoice")
speak.Speak("and the results are")
speak.Speak(results)
