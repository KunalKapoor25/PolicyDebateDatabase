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
driver.get(URL)
for i in range(0, len(sheet)):
    if sheet[i][7] != '':
        continue
    elem = driver.find_element_by_name("search_first")
    s = sheet[i][0]
    s = s[0]
    elem.send_keys(s)
    elem = driver.find_element_by_name("search_last")
    s = sheet[i][1]
    s = s[0]
    elem.send_keys(s)
    elem.send_keys(Keys.ENTER)
    sheet[i][0] = sheet[i][0].strip()
    sheet[i][1] = sheet[i][1].strip()
    print(sheet[i][0] + " " + sheet[i][1])
    all_spans = driver.find_elements_by_xpath("//tr[@role='row']")
    try:
        for span in all_spans:
            m = span.get_attribute("innerHTML")
            if sheet[i][0].lower() in m.lower() and sheet[i][1].lower() in m.lower():
                time.sleep(4)
                link = "https://www.tabroom.com/index/" + m[m.find("href")+6:len(m)-78]
                print(link)
                dataSheet_instance.update_cell(i+2, 8, link)
                break
    except Exception as e:
        print(str(e))
        continue
driver.quit()
    
    
