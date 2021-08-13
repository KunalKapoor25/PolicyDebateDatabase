from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from FileCollection import *
import re

def compileSpeaks(link, first, last):
    spans = driver.find_elements_by_xpath("//a[@class='white padtop padbottom']")
    nums = []
    for i in range(0, len(spans)):
        try:
            text = spans[i].get_attribute('innerHTML')
            if first.lower() in text.lower() and last.lower() in text.lower():
                spans[i].click()
                time.sleep(0.5)
                elem = driver.find_element_by_xpath("//tbody")
                t = elem.get_attribute('innerHTML')
                prenums = re.findall('\d*\.?\d+', t)
                for k in range(0, len(prenums)):
                     if float(prenums[k]) > 20 and float(prenums[k]) <= 30:
                         nums.append(float(prenums[k]))
                return nums
        except Exception as t:
            print(str(t))
            continue
    return nums

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
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[0])
for i in range(1067, len(sheet)):
    try:
        all_nums = []
        count = 0
        if sheet[i][7] == '' or sheet[i][5] != '':
            continue
        driver.get(sheet[i][7])
        all_spans = driver.find_elements_by_xpath("//td[@class='nospace']")
        print(len(all_spans))
        print(sheet[i][0] + " " + sheet[i][1])
        for j in range(0, len(all_spans)):
            if j%8 < 7:
                continue
            if j > 250:
                break
            else:
                m = all_spans[j].get_attribute('innerHTML')
                link = "https://www.tabroom.com" + m[m.find("href")+6:m.find(">")-1]
                driver.switch_to.window(driver.window_handles[1])
                driver.get(link)
                all_nums = all_nums + compileSpeaks(link, sheet[i][0].strip(), sheet[i][1].strip())
            driver.switch_to.window(driver.window_handles[0])
        if len(all_nums) == 0:
            continue
        avg = sum(all_nums)/len(all_nums)
        print(avg)
        dataSheet_instance.update_cell(i+2, 6, avg)
    except Exception as e:
        print(str(e))

driver.quit()
