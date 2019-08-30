# -*- coding: utf-8 -*-
"""
Created on Mon Jan 21 09:29:25 2019

@author: Ajit
"""
from bs4 import BeautifulSoup
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import re
import datetime
today=datetime.datetime.today().strftime('%d_%m_%Y')
file=open(today+'.txt','w')
dic={"Nifty 50":"https://www.investing.com/indices/s-p-cnx-nifty",
     "Nifty Auto":"https://www.investing.com/indices/cnx-auto",
     "Nifty Bank":"https://www.investing.com/indices/bank-nifty",
     "Nifty Commodities":"https://www.investing.com/indices/cnx-commodities",
     "Nifty Energy":"https://www.investing.com/indices/cnx-energy",
     "Nifty Financial Services":"https://www.investing.com/indices/cnx-finance",
     "Nifty FMCG":"https://www.investing.com/indices/cnx-fmcg",
     "Nifty India Consumption":"https://www.investing.com/indices/cnx-consumption",
     "Nifty Infrastructure":"https://www.investing.com/indices/cnx-infrastructure",
     "Nifty IT":"https://www.investing.com/indices/cnx-it",
     "Nifty Media":"https://www.investing.com/indices/cnx-media",
     "Nifty Metal":"https://www.investing.com/indices/cnx-metal",
     "Nifty MNC":"https://www.investing.com/indices/cnx-mnc",
     "Nifty Pharma":"https://www.investing.com/indices/cnx-pharma",
     "Nifty PSU Bank":"https://www.investing.com/indices/cnx-psu-bank",
     "Nifty Realty":"https://www.investing.com/indices/cnx-realty",
     "Nifty Services Sector":"https://www.investing.com/indices/cnx-service-sector"
     }
sector_links=[]
historical_data="-historical-data"
componenets="-components"
base="https://www.investing.com"
driver=webdriver.Chrome()
links_equities=[]
for i in dic.keys():
    sector_links.append(dic[i])
for links in sector_links:
    driver.get(links+historical_data)
    res=driver.execute_script('return document.documentElement.outerHTML')
    soup=BeautifulSoup(res,'lxml')
    date=soup.find('table',{'id':'curr_table'})
    date1=date.find_all('td')
    for name,link in dic.items():
        if links==link:
            print(name+":"+date1[1].text)
            file.write(name+":"+date1[1].text+"\n")
    #print(date1[1].text)
    driver.get(links+componenets)
    res=driver.execute_script('return document.documentElement.outerHTML')
    soup=BeautifulSoup(res,'lxml')
    tab=soup.find('table',{'id':'cr1'})
    tab2=tab.find_all('a')
    for items in tab2:
        links_equities.append(base+items['href'])
        
#ITERATING THROUGH EACH EQUITY
for items in links_equities:
    driver.get(items+historical_data)
    res=driver.execute_script('return document.documentElement.outerHTML')
    soup=BeautifulSoup(res,'lxml')
    div=soup.find('div',{'class':'instrumentHead'})
    name=div.find('h1')
    name=str(div.text)
    name1=re.findall(r'\(.*\)',name)
    name=name1[0]
    name=name.replace('(','')
    name=name.replace(')','')#PRINT THIS
    table=soup.find('table',{'id':'curr_table'})
    td=table.find_all('td')
    value=td[1].text
    print(name+":"+value)
    file.write(name+":"+value+"\n")
    
    
file.close()
driver.close()

    
    
    
    
    
