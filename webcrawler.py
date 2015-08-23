__author__ = 'Shiva'

import requests
from bs4 import BeautifulSoup
import csv

class webcrawl:

    def url_call(self,url='None',baseurl='None', title = 'None'):
        base_url = baseurl
        _url = url
        _title = title
        source_code = requests.get(url)
        if _url == base_url:
            plain_text = source_code.text
            webcrawl.webrequest(self, plain_text)
        elif source_code.status_code == 200:
            if _title == 'None':
                _title = "No title"
            text = [_title,_url,"worked"]
            webcrawl.write_results(self,text)
        else:
            print('broken link')

    def webrequest(self,plain_text):
        soup = BeautifulSoup(plain_text, "html.parser")
        webcrawl.parse_all_links(self,soup)

    def parse_all_links(self, soup):
        #_url = url
        count = 0
        for link in soup.find_all('a'):
            title = link.string
            count +=1
            if(link.get('href')):
                href = link.get('href')
                if href.startswith('http') or href.startswith('https'):
                    #print(href)
                    webcrawl.url_call(self, href,title)
                elif href =='#':
                    #print('No link present')
                    pass
                elif href =='/':
                    pass
                else:
                    href = baseurl + href
                    #print(href)
                    webcrawl.url_call(self, href, title)

                if title=='Subscribe':
                   break
        print(count)

    def write_results(self,text):
        c= csv.writer(open('report.csv','a',newline=''))
        c.writerow(text)

if __name__ == '__main__':
    wc = webcrawl()
    with open('links.csv', newline='',encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            print(row[0])
            baseurl = row[0]
            wc.url_call(baseurl ,baseurl)