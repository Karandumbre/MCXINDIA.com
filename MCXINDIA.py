#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 23 11:15:05 2018

@author: karan.ganesh.dumbre
"""

import pandas as pd
from pandas import ExcelWriter
import re
from bs4 import BeautifulSoup
from selenium import webdriver
import time
from selenium.webdriver.support.select import Select
from datetime import datetime


driver = webdriver.Chrome(executable_path='/Users/karan.ganesh.dumbre/chromedriver');
driver.get("https://www.mcxindia.com/market-data/historical-data#")

Datewise = driver.find_element_by_id('Datewise')
Datewise.click();

Select1 = Select(driver.find_element_by_id("ddlInstrumentName"))
Select1.select_by_visible_text("FUTCOM")


Select2 = Select(driver.find_element_by_id("cph_InnerContainerRight_C004_ddlSegment"))
Select2.select_by_visible_text("BULLION")


Select3 = Select(driver.find_element_by_id("ddlCommodityHead"))
Select3.select_by_visible_text("GOLD")

Select4 = Select(driver.find_element_by_id("ddlCommodityContract"))
Select4.select_by_visible_text("GOLD")

def date(str):
    driver.find_element_by_id(str).click()
    selectYear = Select(driver.find_element_by_css_selector('body > div.datepick-popup > div > div.datepick-month-row > div > div > select:nth-child(2)'))
    selectYear.select_by_visible_text("2018")
    time.sleep(1)
    
    selectMonth = Select(driver.find_element_by_css_selector('body > div.datepick-popup > div > div.datepick-month-row > div > div > select:nth-child(1)'))
    selectMonth.select_by_visible_text("April")
    time.sleep(1)
    
    colName = driver.find_element_by_css_selector("[title=\"Select Monday, Apr 16, 2018\"]")
    colName.click();
    time.sleep(1)

date('txtFromDate')
date('txtToDate')

driver.find_element_by_id("btnSummary").click()
time.sleep(4) ### delay to load the data in chrome
    
html = driver.page_source
soup = BeautifulSoup(html)

values = []
heading = []

row = soup.find('div',{'class':'main-row'})

####Extract Data

def extract(a):
    col = row.find('div',{'class':a})

    col_head = col.find_all('div',{'class':'col-head'})
    
    for i in col_head:
        i = i.text.encode('ascii','ignore')
        i = i.strip();
        heading.append(i.replace('\n',''))
    
    val=col.find_all('div',{'class':''})
    
    for i in val:
        values.append(i.text.encode('ascii','ignore'))

##### For values in col head 1
extract('col-1')

##### For values in col head 2
extract('col-2')
count = 0
for i in heading:
    
    if(i == 'Instrument'):
        Instrument = values[count]
    if(i == 'Date'):
        Date = values[count]
    if(i == 'Traded Contract (Lots)'):
        Traded_Contract = values[count]
    if(i == "Quantity                    (000's)"):
        Quantity = values[count]
    if(i == 'Total Value (Lacs)'):
        Total_value = values[count]
    count = count + 1

### Ectract the quantity data
Quantity = re.findall(r'\d+', Quantity)[0]

Quantity = int(Quantity)
Total_value = float(Total_value)


#### Creating a dataframe

df = pd.DataFrame({
       "Instrument":Instrument,
       "Date":Date,
       "Traded_Contract":Traded_Contract,
       "Quantity":Quantity,
       "Total_value":Total_value
        }, index=[0])

#Export To excel
writer = ExcelWriter('MCXINDIA.xlsx')
df.to_excel(writer,'Sheet1')


Date = datetime.strptime(Date, '%d %b %Y')
Date = Date.strftime('%d %m %Y')

Price = (Total_value * 100000)/(8317 * 100)

df2 = pd.DataFrame({
       "Date":Date,
       "Price":Price
        }, index=[0])

#Export To excel
df2.to_excel(writer,'Sheet2')
writer.save() 