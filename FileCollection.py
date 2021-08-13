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
