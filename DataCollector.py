import gspread
import pandas as pd
import time
from oauth2client.service_account import ServiceAccountCredentials

def removeSyntax(s):
    val = s.count('\'')
    if val < 2:
        return s
    return s[s.find('\'')+1:s.rfind('\'')]+s[s.rfind('\'')]

def sourceSheetCompile1(n):
    print(1)
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/williamyang/Desktop/001 - WMDs/PrefBriefs/My First Project-3c36563e88f3.json', scope)
    client = gspread.authorize(creds)
    sourceSheet = client.open('KT - Prefs')
    sourceSheet_instance = sourceSheet.get_worksheet(n)
    firstNames = sourceSheet_instance.col_values(1)
    for i in range(0, len(firstNames)):
        firstNames[i] = removeSyntax(str(firstNames[i]))
    firstNames.pop(0)
    lastNames = sourceSheet_instance.col_values(2)
    for i in range(0, len(lastNames)):
        lastNames[i] = removeSyntax(str(lastNames[i]))
    lastNames.pop(0)
    percent = sourceSheet_instance.col_values(3)
    percent.pop(0)
    for i in range(0, len(percent)):
        percent[i] = float(removeSyntax(str(percent[i])))
    names = list(zip(firstNames, lastNames, percent))
    return names

def sourceSheetCompile2(n):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/williamyang/Desktop/001 - WMDs/PrefBriefs/My First Project-3c36563e88f3.json', scope)
    client = gspread.authorize(creds)

    sourceSheet = client.open('KT - Prefs')
    sourceSheet_instance = sourceSheet.get_worksheet(n)
    names = sourceSheet_instance.col_values(1)
    for i in range(0, len(names)):
        names[i] = removeSyntax(str(names[i]))
    names.pop(0)
    firstNames = []
    lastNames = []
    for i in range(0, len(names)):
        s = names[i]
        lastNames.append(s[0:s.find(',')])
        firstNames.append(s[s.find(',')+2:len(s)])
    percent = sourceSheet_instance.col_values(3)
    percent.pop(0)
    for i in range(0, len(percent)):
        percent[i] = float(removeSyntax(str(percent[i])))
    names = list(zip(firstNames, lastNames, percent))
    return names

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

def update(source, data):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/williamyang/Desktop/001 - WMDs/PrefBriefs/My First Project-3c36563e88f3.json', scope)
    client = gspread.authorize(creds)
    dataSheet = client.open('Master Data Sheet')
    dataSheet_instance = dataSheet.get_worksheet(1)
    for i in range(0, len(source)):
        time.sleep(4)
        print(source[i][1] + " " + source[i][0])
        for j in range(0, len(data)):
            if source[i][0].lower() == data[j][0].lower() and source[i][1].lower() == data[j][1].lower():
                val = (data[j][2]*data[j][3]+source[i][2])/(data[j][3]+1)
                dataSheet_instance.update_cell(j+2, 4, data[j][3]+1)
                dataSheet_instance.update_cell(j+2, 3, val)
                break
            if j == len(data)-1:
                addList = []
                addList.append(source[i][0])
                addList.append(source[i][1])
                addList.append(source[i][2])
                addList.append(1)
                data.append(addList)
                dataSheet_instance.update_cell(len(data)+1, 1, addList[0])
                dataSheet_instance.update_cell(len(data)+1, 2, addList[1])
                dataSheet_instance.update_cell(len(data)+1, 3, addList[2])
                dataSheet_instance.update_cell(len(data)+1, 4, addList[3])
            

for i in range(1, 12):
    print(i)
    choice = 1
    if choice == 1:
        source = sourceSheetCompile1(i)
    else:
        source = sourceSheetCompile2(i)
    data = dataSheetCompile(1)
    print(len(source))
    print(len(data))
    update(source, data)


