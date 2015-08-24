__author__ = 'Shiva'
import requests
import re

url = "http://www.gewater.com/"
#connect to a URL
website = requests.get(url)

#read html code
html = website.text

#use re.findall to get all the links
links = re.findall('"((http|ftp)s?://.*?)"', html)

for link in links:
    print(link)