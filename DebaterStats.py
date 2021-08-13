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

#add in program that scans repeat names and deletes them if they don't match the exact text on the google sheet
def addNames(f, l):
    FILE = "names.txt"
    fn = open(FILE, 'r')
    T = fn.readlines()
    fn.close()
    print(len(T))
    for j in range(0, len(T)):
        T[j] = T[j].strip().lower()
    for i in range(0, len(f)):
        T.append((f[i] + "()" + l[i]).lower())
    T = list(set(T))
    print(len(T))
    fw = open(FILE, 'w')
    for k in range(0, len(T)):
        if ">" in T[k]:
            continue
        fw.write(T[k] + "\n")
    fw.close()
    
def extractFile():
    FILE = "names.txt"
    f = open(FILE, 'r')
    T = f.readlines()
    f.close()
    for j in range(0, len(T)):
        T[j] = T[j].strip()
    return T
    
def extractNames():
    running = True
    mainURL = "https://opencaselist.paperlessdebate.com/"
    driver.get(mainURL)
    time.sleep(1)
    f_list = []
    l_list = []
    count = 0
    while running:
        count += 1
        print(count)
        try:
            first = driver.find_elements_by_xpath("//td[@class='first_name link typetext']")
            last = driver.find_elements_by_xpath("//td[@class='last_name link typetext']")
            for span in first:
                f = span.get_attribute('innerHTML').strip().lower()
                f_list.append(f)
            for sp in last:
                l = sp.get_attribute('innerHTML').strip().lower()
                l_list.append(l)
            check = driver.find_element_by_xpath("//span[@class='controlPagination']")
            text = check.get_attribute('innerHTML')
            if "noNextPagination" in text:
                break
            else:
                driver.find_element_by_xpath("//a[@title='Next Page']").click()
                time.sleep(1)
        except Exception as e:
            print(str(e))
    addNames(f_list, l_list)

def extractNames2():
    running = True
    mainURL = "https://opencaselist16.paperlessdebate.com"
    driver.get(mainURL)
    driver.execute_script('''window.open();''')
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    f_list = []
    l_list = []
    spans = driver.find_elements_by_xpath("//span[@class='wikilink']")
    for i in range(6, len(spans)):
        elem_list = []
        try:
            s = spans[i]
            t = s.get_attribute('innerHTML')
            link = mainURL + t[t.find("href")+6:t.find(">")-1]
            driver.switch_to.window(driver.window_handles[1])
            driver.get(link)
            name = driver.find_element_by_id("document-title").get_attribute('innerHTML')
            name = name[name.find("<h1>")+4:name.find("</h1>")]
            print(name)
            table_text = driver.find_element_by_xpath("//table[@class='grid sortable doOddEven']").get_attribute('innerHTML')
            for j in range(len(table_text)-len(name)):
                if table_text[j:j+len(name)] == name:
                    sub = table_text[j:len(table_text)].find("</p>")+j
                    sub_text = table_text[j+len(name)+1:sub]
                    elem_list.append(sub_text)            
            driver.switch_to.window(driver.window_handles[0])
            for k in range(0, len(elem_list)):  
                names = elem_list[k].split()
                if len(names) == 5:
                    f_list.append(names[0])
                    l_list.append(names[1])
                    f_list.append(names[3])
                    l_list.append(names[4])        
        except:
            elem_list = []
            continue
    for m in range(0, len(l_list)):
        if m == len(l_list)-2:  
            if "type=" in l_list[m] or "\">" in l_list[m]:
                del l_list[m]
                del f_list[m]
            break
        else:
            if "type=" in l_list[m] or "\">" in l_list[m]:
                del l_list[m]
                del f_list[m]
    addNames(f_list, l_list)
    
    
    

def checkName(f, l):
    for i in range(0, len(sheet)):
        if f.lower() == sheet[i][0].lower() and l.lower() == sheet[i][1].lower():
            return True
    return False
    
def addNameStats():
    names = extractFile()
    URL = "https://www.debatestat.com/ind/"
    for i in range(6008, len(names)):
        print(i)
        driver.get(URL)
        elem = driver.find_element_by_id("dInput")
        n = names[i]
        n = n[0:n.find('(')] + " " + n[n.find(')')+1:len(n)]
        f = n[0:n.find(" ")]
        l = n[n.find(" ")+1:len(n)]
        if checkName(f, l):
            continue
        elem.send_keys(n)
        driver.find_element_by_xpath("//h1[@class='title is-2']").click()
        driver.find_element_by_xpath("//input[@class='button is-link']").click()
        cur = driver.current_url
        try:
            print(n)
            if "tabroom" in cur:
                ID = cur[cur.find("id2=")+4:len(cur)]
                print(ID)
                sheet.append([f, l, ID, "", "", "", ""])
                dataSheet_instance.update_cell(len(sheet)+1, 1, f)
                dataSheet_instance.update_cell(len(sheet)+1, 2, l)
                dataSheet_instance.update_cell(len(sheet)+1, 3, ID)            
            else:
                elem = driver.find_elements_by_xpath("//a[@href]")
                lnks = []
                for span in elem:
                    lnks.append(span.get_attribute("href"))
                for j in range(6, len(lnks)):
                    driver.get(lnks[j])
                    cur = driver.current_url
                    ID = cur[cur.find("id2=")+4:len(cur)]
                    print(ID)
                    w = driver.find_element_by_xpath("//h4[@class='nospace']")
                    text = w.get_attribute("innerHTML")
                    t1 = text.split()
                    sheet.append([t1[0], t1[1], ID, "", "", "", ""])
                    dataSheet_instance.update_cell(len(sheet)+1, 1, t1[0].lower())
                    dataSheet_instance.update_cell(len(sheet)+1, 2, t1[1].lower())
                    dataSheet_instance.update_cell(len(sheet)+1, 3, ID)   
                    

        except Exception as e:
            print(str(e))

    
def calcSpeaks():
    for i in range(0, len(sheet)):
        prenums = []
        nums = []
        if sheet[i][3] != "" and sheet[i][4] != "" and sheet[i][6] != "":
            continue
        try:
            print(i)
            print(sheet[i][0] + " " + sheet[i][1])
            time.sleep(3)
            URL = "https://www.tabroom.com/index/results/team_lifetime_record.mhtml?id1=&id2=" + str(sheet[i][2])
            driver.get(URL)
            elem = driver.find_element_by_xpath("//div[@class='nospace padtop']")
            t = elem.get_attribute('innerHTML')
            prenums = re.findall('\d*\.?\d+', t)
            for k in range(0, len(prenums)):
                if float(prenums[k]) > 26 and float(prenums[k]) <= 30:
                    nums.append(float(prenums[k]))
            if len(nums) == 0:
                continue
            elif len(nums) >= 30:
                s = 0
                for j in range(len(nums)-30, len(nums)):
                    s += nums[j]
                lastAvg = s/30
                dataSheet_instance.update_cell(i+2, 4, lastAvg)
                dataSheet_instance.update_cell(i+2, 5, sum(nums)/len(nums))
                dataSheet_instance.update_cell(i+2, 7, len(nums))
            else:
                dataSheet_instance.update_cell(i+2, 4, sum(nums)/len(nums))
                dataSheet_instance.update_cell(i+2, 5, sum(nums)/len(nums))
                dataSheet_instance.update_cell(i+2, 7, len(nums))
        except Exception as e:
            print(str(e))
            
def findInstances(t, s):
    running = True
    ind = 0
    tot = 0
    e = []
    while running:
        ind = t.find(s)
        if ind < 0:
            break
        else:
            tot += ind
            ind += len(s)
            e.append(tot)
            tot += len(s)
            t = t[ind:len(t)]
    return e
        
    
def calcWinRate():
    #need to solve issue of repeating ID values
    for i in range(0, len(sheet)):
        if sheet[i][5] != '':
            continue
        try:
            time.sleep(1)
            print(i)
            print(sheet[i][0] + " " + sheet[i][1])
            URL = "https://www.tabroom.com/index/results/team_lifetime_record.mhtml?id1=&id2=" + str(sheet[i][2])
            driver.get(URL)
            elem = driver.find_element_by_id("results_table")
            t = elem.get_attribute('innerHTML')
            td = findInstances(t, "<td>")
            td2 = findInstances(t, "</td>")
            entries = []
            for j in range(0, len(td)):
                entries.append(t[td[j]+4:td2[j]])
            pr = entries[11]
            pr = str(pr[0:pr.find("\n")])
            dataSheet_instance.update_cell(i+2, 6, pr)
        except Exception as e:
            print("ERROR")
            print(str(e))

def getNames():
    #this method is way too slow I have a new algorithm for filling in gaps of names
    #extract names from wiki page not from user-index but from team pages
    #then loop through everything and check whether name at certain index in the i for loop is on the wiki-page names list
    col = dataSheet_instance.col_values(3)
    policy_names = extractFile()
    for i in range(106700, 999999):
        if str(i) in col:
            continue
        URL = "https://www.tabroom.com/index/results/team_lifetime_record.mhtml?id1=&id2=" + str(i)
        driver.get(URL)
        try:    
            elem = driver.find_element_by_xpath("//h4[@class='martopmore']")
            t = elem.get_attribute('innerHTML').split(" ")
            s = t[2]
            s1 = s[0:s.find("\n\t\t\t\t\t\t")]
            s2 = s[s.find("\n\t\t\t\t\t\t"):len(s)].strip()
            m = (s1 + "()" + s2).lower()
            if m not in policy_names:
                continue
            print(i)
            print(s1)
            print(s2)
            sheet.append([s1, s2, i, "", "", "", ""])
            dataSheet_instance.update_cell(len(sheet)+1, 3, i)
            dataSheet_instance.update_cell(len(sheet)+1, 1, s1)
            dataSheet_instance.update_cell(len(sheet)+1, 2, s2)
            time.sleep(2.5)
        except:
            continue
        
        
            
CHROME_DRIVER = "/Users/williamyang/Desktop/001 - WMDs/driver/chromedriver"
options = webdriver.ChromeOptions()
options.add_argument('headless')
driver = webdriver.Chrome(executable_path=CHROME_DRIVER, options=options)
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/williamyang/Desktop/001 - WMDs/PrefBriefs/My First Project-3c36563e88f3.json', scope)
client = gspread.authorize(creds)
dataSheet = client.open('Master Data Sheet')
dataSheet_instance = dataSheet.get_worksheet(2)
sheet = dataSheetCompile(2)


#extractNames()
#extractNames2()
#addNameStats()
#calcSpeaks()
calcWinRate()
#testUpdate()
#getNames()
driver.quit()
