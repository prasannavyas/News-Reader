import requests
import json
from win32com.client import Dispatch
def speak3(str):
    Dispatch ("SAPI.SpVoice")
    #speak3(str)

for i in range(10):
    speak3("Today s headlines are")
url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=f4b478647d3147ecbd698ae73f88f823"
news = requests.get(url)
news_dict = news.json()
arts = news_dict['articles']
for article in arts:
    speak3(article['title'])
    print(article['title'])
    speak3("Next news is    ")

speak3("Thats all for today. See you tommorow")