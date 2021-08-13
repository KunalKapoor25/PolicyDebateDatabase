from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

def removeSyntax(s):
    val = s.count('\'')
    if val < 2:
        return s
    return s[s.find('\'')+1:s.rfind('\'')]+s[s.rfind('\'')]

def dataSheetCompile(n):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/williamyang/Desktop/001 - WMDs/PrefBriefs/My First Project-3c36563e88f3.json', scope)
    client = gspread.authorize(creds)
    dataSheet = client.open('Master Data Sheet')
    dataSheet_instance = dataSheet.get_worksheet(n)
    sheet = dataSheet_instance.get_all_values()
    sheet.pop(0)
    return sheet

def numJudges():
    span = driver.find_element_by_xpath("//span[@class='threequarters nospace']")
    s = span.get_attribute('innerHTML')
    s = s[s.find(":")+2:s.find("</h4>")]
    return int(s)    
    
CHROME_DRIVER = "/Users/williamyang/Desktop/001 - WMDs/driver/chromedriver"
options = webdriver.ChromeOptions()
options.add_argument('headless')
driver = webdriver.Chrome(executable_path=CHROME_DRIVER, options=options)
URL = "https://www.tabroom.com/index/paradigm.mhtml"
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/williamyang/Desktop/001 - WMDs/PrefBriefs/My First Project-3c36563e88f3.json', scope)
client = gspread.authorize(creds)
dataSheet = client.open('Master Data Sheet')
dataSheet_instance = dataSheet.get_worksheet(1)
sheet = dataSheetCompile(1)
status = 1
driver.get(URL)
for i in range(0, len(sheet)):
    if sheet[i][7] != "" and sheet[i][4] == "":
        driver.get(sheet[i][7])
        time.sleep(1)
        try:
            print(sheet[i][0] + " " + sheet[i][1])
            all_spans = driver.find_elements_by_xpath("//tr[@role='row']")
            count = len(all_spans)-2
            time.sleep(3)
            dataSheet_instance.update_cell(i+2, 5, count)               
        except Exception as e:
            print(str(e))
            continue
    else:
        continue
    
    
driver.quit()
