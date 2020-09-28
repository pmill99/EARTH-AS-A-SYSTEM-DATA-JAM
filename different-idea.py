from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import Select
from datetime import timedelta, date
from time import sleep
import openpyxl

current_date = date.today()
current_month = current_date.strftime('%m')
current_day = current_date.strftime('%d')

datafile = openpyxl.load_workbook('datasheet.xlsx')
sheet = datafile.active

start_data_collection = date(2020, 9, 1)
end_data_collection = date(2020, 10, 31)

delta = timedelta(days=1)

datelist = []
monthlist = []
daylist = []

while start_data_collection <= end_data_collection:
    month = start_data_collection.strftime('%m')
    day = (start_data_collection.strftime('%d'))
    today = month + '-' + day
    datelist.append(today)
    monthlist.append(month)
    daylist.append(day)
    start_data_collection += delta

del start_data_collection
del end_data_collection
del delta
del month
del day
del today
del current_date

row = 3

for i in range(len(datelist)):
    #Open browser
    browser = webdriver.Chrome()
    browser.get(f'https://www.weatherforyou.com/reports/index.php?forecast=pass&pass=archive&zipcode=13165&pands=13165&place=waterloo&state=ny&icao=KPEO&country=us&month={monthlist[i]}&day={daylist[i]}&year=2020&dosubmit=Go')
    sleep(3)
    
    #Finding elements by Xpath
    hightemperature =  browser.find_element_by_xpath('//*[@id="middlepagecontent"]/table/tbody/tr/td[1]/table[2]/tbody/tr[3]/td[2]')
    lowtemperature = browser.find_element_by_xpath('//*[@id="middlepagecontent"]/table/tbody/tr/td[1]/table[2]/tbody/tr[3]/td[3]')
    avgtemperature = browser.find_element_by_xpath('//*[@id="middlepagecontent"]/table/tbody/tr/td[1]/table[2]/tbody/tr[3]/td[4]')
    dewpoint = browser.find_element_by_xpath('//*[@id="middlepagecontent"]/table/tbody/tr/td[1]/table[2]/tbody/tr[4]/td[4]')
    if monthlist[i] == current_month and daylist[i] <= current_day:
        precip = browser.find_element_by_xpath('//*[@id="middlepagecontent"]/table/tbody/tr/td[1]/table[2]/tbody/tr[9]/td[4]')
    sleep(3)

    #Writing to excel file
    sheet[f'A{row}'] = datelist[i]              #Date
    sheet[f'B{row}'] = hightemperature.text     #High Temperature
    sheet[f'C{row}'] = lowtemperature.text      #Low Temperature
    sheet[f'D{row}'] = avgtemperature.text      #Avg Temperature
    sheet[f'E{row}'] = dewpoint.text            #Dewpoint
    if monthlist[i] == current_month and daylist[i] <= current_day:
        sheet[f'F{row}'] = precip.text          #Precipitation

    #Save Excel spreadsheet and increment row    
    datafile.save('datasheet.xlsx')
    row += 1

    #Close browser
    browser.close()
    sleep(3)

#Save and Exit Excel
datafile.save()
datafile.close()