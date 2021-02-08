# -*- coding: utf-8 -*-
"""
Created on Fri Jan 29 12:38:20 2021

@author: jasmin
"""
import requests
from bs4 import BeautifulSoup
import time
import urllib.parse
import zipfile
import io

def get_url_content(url):
    time.sleep(.1)
    return requests.get(url).text

url = 'https://statistik.arbeitsagentur.de/SiteGlobals/Forms/Suche/Einzelheftsuche_Formular.html?submit=Suchen&topic_f=gemeinde-arbeitslose-quoten'

text = get_url_content(url)

soup = BeautifulSoup(text, "html.parser")
review_content = soup.find_all('div', {'class': 'teaser type-1 StatisticDocument row'})
tags = review_content[1].find(href=True)
link = tags.get('href')

full_url = urllib.parse.urljoin(url, link)

r = requests.get(full_url, stream=True)
z = zipfile.ZipFile(io.BytesIO(r.content))
f = z.extractall()
   


