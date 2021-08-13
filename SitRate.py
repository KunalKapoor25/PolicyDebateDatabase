from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from FileCollection import *

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
    if sheet[i][7] == '' or sheet[i][6] != '':
        continue
    driver.get(sheet[i][7])
    all_spans = driver.find_elements_by_xpath("//td[@class='nowrap']")
    decision = ''
    panel_decision = ''
    sit = 0
    count = 0
    for j in range(0, len(all_spans)):
        if j%3 == 0:
            continue
        elif j%3 == 1:
            decision = all_spans[j].get_attribute('innerHTML').strip()
        elif j%3 == 2:
            panel_decision = all_spans[j].get_attribute('innerHTML').strip()
            if panel_decision == '':
                decision = ''
                continue
            else:
                count += 1
                if decision not in panel_decision:
                    sit += 1
    if count == 0:
        continue
    percent = sit/count*100
    print(str(percent) + "%")
    dataSheet_instance.update_cell(i+2, 7, str(percent)+"%")
driver.quit()          
        
