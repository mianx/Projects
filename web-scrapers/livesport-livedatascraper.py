import requests
from bs4 import BeautifulSoup
from urllib.request import urlopen
import pandas as pd
from selenium import webdriver
import time
import numpy as np

urls = ['https://www.livesport.com/nl/handbal/nederland/handbalnl-league/schema/',
       'https://www.livesport.com/nl/handbal/nederland/eredivisie/schema/',
       'https://www.livesport.com/nl/handbal/nederland/eredivisie-vrouwen/schema/',
       'https://www.livesport.com/nl/hockey/nederland/hoofdklasse/schema/',
        'https://www.livesport.com/nl/hockey/nederland/hoofdklasse-vrouwen/schema/',
       'https://www.livesport.com/nl/futsal/nederland/eredivisie/schema/',
       'https://www.livesport.com/nl/american-football/nederland/eredivisie/schema/',
       'https://www.livesport.com/nl/rugby/nederland/ereklasse/schema/',
        'https://www.livesport.com/nl/volleybal/nederland/eredivisie/schema/',
        'https://www.livesport.com/nl/volleybal/nederland/eredivisie-vrouwen/schema/']

driver = webdriver.Chrome(r"C:\SeleniumDrivers\chromedriver.exe")
for url in urls:
    driver.get(url)
    html = driver.page_source
    time.sleep(2)
    soup = BeautifulSoup(html)
    data = []
    date_time = soup.find_all('div',{'class','event__time'})
    home_team = soup.find_all('div',{'class','event__participant event__participant--home'})
    away_team = soup.find_all('div',{'class','event__participant event__participant--away'})
    for dt,home,away in zip(date_time,home_team,away_team):
        temp = dt.text.split()
        
        dic = {
            " " : np.nan,
            'DATE' : '2022' + '-'+ temp[0][:-1].split('.')[1] + '-'+ temp[0][:-1].split('.')[0],
            'TIME' : temp[1][:-2],
            'HOME TEAM' : home.text,
            'AWAY TEAM' : away.text
        }
        data.append(dic)
    df = pd.DataFrame(data)
    df.set_index(df.columns[0],inplace= True)
    title = soup.title.text.split() 
    title[0] = title[0].replace(':','')
    #temp = ' '.join(title[:-4]) + ".csv"
    temp = 'American football.csv'
    df.to_csv(temp,sep=';')

driver.quit()