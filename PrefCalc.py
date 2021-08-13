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
    firstNames = dataSheet_instance.col_values(1)
    for i in range(0, len(firstNames)):
        firstNames[i] = removeSyntax(str(firstNames[i]))
    firstNames.pop(0)
    lastNames = dataSheet_instance.col_values(2)
    for i in range(0, len(lastNames)):
        lastNames[i] = removeSyntax(str(lastNames[i]))
    lastNames.pop(0)
    percent = dataSheet_instance.col_values(3)
    percent.pop(0)
    for i in range(0, len(percent)):
        percent[i] = float(removeSyntax(str(percent[i])))
    count = dataSheet_instance.col_values(4)
    count.pop(0)
    for i in range(0, len(count)):
        count[i] = float(removeSyntax(str(count[i])))
    names = list(zip(firstNames, lastNames, percent, count))
    return names

def numJudges():
    span = driver.find_element_by_xpath("//span[@class='threequarters nospace']")
    s = span.get_attribute('innerHTML')
    s = s[s.find(":")+2:s.find("</h4>")]
    return int(s)

def gather(sheets, names, num_judges):
    rank = []
    if len(names) > num_judges*2:
        for i in range(0, len(names), 3):
            temp = []
            first = names[i]
            last = names[i+2]
            temp.append(first)
            temp.append(last)
            for j in range(0, len(sheet)):
                if first.lower() == sheet[j][0].lower() and last.lower() == sheet[j][1].lower():
                    if len(temp) > 2:
                        continue
                    temp.insert(0, sheet[j][2])
            if len(temp) == 2:
                temp.insert(0, 999)
            rank.append(temp)
    else:  
        for i in range(0, len(names), 2):
            temp = []
            first = names[i]
            last = names[i+1]
            temp.append(first)
            temp.append(last)
            for j in range(0, len(sheet)):
                if first.lower() == sheet[j][0].lower() and last.lower() == sheet[j][1].lower():
                    if len(temp) > 2:
                        continue
                    temp.insert(0, sheet[j][2])
            if len(temp) == 2:
                temp.insert(0, 999)
            rank.append(temp)
    return rank
    
print("Input Judge Page Link")
URL = input()
print("Policy (1) or K (0)?")
choice = int(input())
CHROME_DRIVER = "/Users/williamyang/Desktop/001 - WMDs/driver/chromedriver"
options = webdriver.ChromeOptions()
options.add_argument('headless')
driver = webdriver.Chrome(executable_path=CHROME_DRIVER, options=options)

names = []


driver.get(URL)
all_spans = driver.find_elements_by_xpath("//a[@class='white full padvert']")
count = 0
sheet = dataSheetCompile(choice)
for span in all_spans:
    name = span.get_attribute('innerHTML')
    name = name.strip()
    names.append(name)

num_judges = numJudges()

rank = gather(sheet, names, num_judges)
rank.sort()
print("PREFS GENERATED:")
for i in range(0, len(rank)):
    if rank[i][0] == 999:
        print(str(i+1) + " - " + rank[i][1] + " " + rank[i][2] + " - NO DATA")
    else:   
        print(str(i+1) + " - " + rank[i][1] + " " + rank[i][2])
     
driver.quit()
