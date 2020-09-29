import time, datetime
from datetime import timedelta
import traceback
import pandas as pd
import numpy as np
import xlsxwriter
import os
from IPython.display import display_html
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# VARIABLES

i = 1 # Respect the Almighty Loop Tracker! Praise be to i
account_id = ""
LossID = ""
errException = "Dummy Exception"

# LOG FILE
logName = 'Account List Reformat & Checks.txt'

# SLEEP TIMER
sleepTimer = 2

# OPEN THE ACCOUNT RENEWALS LIST - CHANGE FILE NAME AND SHEET NAME IN THE LINE DIRECTLY BELOW THIS
dfem = pd.read_excel(open('C://Users//alessio//Desktop//Python Scripts//Accounts//Account renewals August 2020.xlsx','rb'), sheet_name='Account renewals August 2020')

# REMOVE DUPLICATES AND SPECIAL CHARACTERS FROM ACCOUNT NAMES
dfem.drop_duplicates(subset ="Customer Name", inplace = True)
dfem['Customer Name'] = dfem['Customer Name'].str.replace('!','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('@','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('#','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('$','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('%','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('"','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('^','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('*','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('~','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('<','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('>','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('/','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('?','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace(';','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace(':','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('|','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('[','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace(']','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('{','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('}','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('+','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('=','')
dfem['Customer Name'] = dfem['Customer Name'].str.replace('*','')

# REMOVE MN CODES LEAVING ONLY FN CODES
dfem = dfem[~dfem['Mnfn Code'].str.startswith('MN')]
dfem = dfem[~dfem['Mnfn Code'].str.startswith('Mn')]
dfem = dfem[~dfem['Mnfn Code'].str.startswith('mN')]
dfem = dfem[~dfem['Mnfn Code'].str.startswith('mn')]

# REMOVE CAPTIVE ACCOUNTS
dfem.drop(dfem[(dfem['Is Captive'] == 'Y') | (dfem['Is Captive'] == 1)].index, inplace=True)

# REMOVE NON-EUROPEAN ACCOUNTS
dfem.drop(dfem[(dfem['Producing Country'] != 'AUSTRIA') & (dfem['Producing Country'] != 'BAHRAIN') & (dfem['Producing Country'] != 'BELGIUM') & (dfem['Producing Country'] != 'CZECH REPUBLIC') & (dfem['Producing Country'] != 'DENMARK') & (dfem['Producing Country'] != 'FINLAND') & (dfem['Producing Country'] != 'FRANCE') & (dfem['Producing Country'] != 'GERMANY') & (dfem['Producing Country'] != 'HUNGARY') & (dfem['Producing Country'] != 'IRELAND') & (dfem['Producing Country'] != 'ITALY') & (dfem['Producing Country'] != 'NETHERLANDS') & (dfem['Producing Country'] != 'NORWAY') & (dfem['Producing Country'] != 'PAKISTAN') & (dfem['Producing Country'] != 'POLAND') & (dfem['Producing Country'] != 'PORTUGAL') & (dfem['Producing Country'] != 'RUSSIA') & (dfem['Producing Country'] != 'SOUTH AFRICA') & (dfem['Producing Country'] != 'SPAIN') & (dfem['Producing Country'] != 'SWEDEN') & (dfem['Producing Country'] != 'SWITZERLAND') & (dfem['Producing Country'] != 'TURKEY') & (dfem['Producing Country'] != 'EGYPT') & (dfem['Producing Country'] != 'UNITED KINGDOM')].index, inplace=True)

# REMOVE DUPLICATE FN NUMBERS
dfem.drop_duplicates(subset ="Mnfn Code", inplace = True)

# REPLACE MULTILINE, ACCIDENT & HEALTH AND TERRORISM UNDER LOB
dfem['Line Of Business'] = dfem['Line Of Business'].str.replace('Multiline','Multiple Lines Of Business')
dfem['Line Of Business'] = dfem['Line Of Business'].str.replace('Terrorism','Multiple Lines Of Business')
dfem['Line Of Business'] = dfem['Line Of Business'].str.replace('Accident & Health','A&H')

# DUPLICATE INCEPTION DATE - The inception date needs to be formatted in two different ways: one for the loss run description (dd mmm yy) and the other for the renewal date on the loss run (mm-dd), meaning two inception date fields are needed
dfem['Renewal Date'] = dfem['Effective Date']

# REFORMAT INCEPTION DATE
dfem['Effective Date'] = dfem['Effective Date'].dt.strftime("%d %b %y")

#REFORMAT RENEWAL DATE
dfem['Renewal Date'] = dfem['Renewal Date'].dt.strftime("%m-%d")

# DROP UNNECCESSARY COLUMNS
dfem.drop(['Country Name','Country Of Risk','Country Of Policy','Producing Region','Producing Company','Business Group','Producing Office Name','Servicing Company','Date Completed','Gross Written Prem Amount','Ace Or Non Affiliate','Gross Written Prem Curr Desc','No Of Polices','Us Value Gwp','Date Due','Service Office Region','Responsible Person','Underwriter','Service Coordinator','Delay Reason','Comment','Market Segment','Isrenewal','REN or NB Merged','Program Type','Is chubbrenewal','Islegacychubbnewbusiness','Policynumber','Action'], axis = 1, inplace = True)

# ACCOUNT CHECK ARRAYS
print('Data loaded...')
mnfnNumbers = dfem['Mnfn Code'].tolist()
fnrenewals = []
accntstatus = []
loglist = []

# LOG SCRIPT START
timeStamp = str(time.ctime()) # Get Current Time
f = open(logName, 'a') # Open Log File a = append mode so it wont overwrite anything if file exists
f.write('\n' + '0----------Script Started at ' + timeStamp) #Log Competion in File
f.close() # Save and Close Log File
print('0----------Script Started at ' + timeStamp)

# BEGIN CHECKING LOOP
# CHECK ACCOUNTS ARE PRESENT IN WEBAPP
for c, value in enumerate(mnfnNumbers):
    #OPEN INTERNET EXPLORER
    chrome = webdriver.Ie() #changed to internet explorer because Chrome does not have all elements showing
    actions = ActionChains(chrome)
    #GO TO MLNR WEB PAGE
    chrome.get('http://webapplink.com/Default.aspx') #goes to the home page
    chrome.get('http://webapplink.com/AccountRef.aspx') #goes to the loss run page
    wait = WebDriverWait(chrome, 5)
    chrome.find_element_by_name('search$fn').send_keys(value) #ENTERS THE FN CODE
    chrome.find_element_by_name('search$button').click() #presses the search button
    try:

#FIND THE ACCOUNT
        try:
            element = wait.until(EC.presence_of_element_located((By.NAME, 'account$edit$button')))
            element = wait.until(EC.element_to_be_clickable((By.NAME, 'account$edit$button')))
            time.sleep(sleepTimer)
#FIND THE LOSS RUN
            chrome.find_element_by_name('account$edit$button').click() #edit account
            time.sleep(sleepTimer)
            try:
                try:
                    element = wait.until(EC.presence_of_element_located((By.ID, 'TabOne')))
                    element = wait.until(EC.element_to_be_clickable((By.ID, 'TabOne')))
                    chrome.find_element_by_id("TabOne").click()
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.ID, 'TabOne')))
                    element = wait.until(EC.element_to_be_clickable((By.ID, 'TabOne')))
                    chrome.find_element_by_id("TabOne").click()
                try:
                    element = wait.until(EC.presence_of_element_located((By.ID, 'LossIDbutton')))
                    element = wait.until(EC.element_to_be_clickable((By.ID, 'LossIDbutton')))
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'LossIDbutton')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'LossIDbutton')))
#WHEN BOTH ACCOUNT AND LOSS RUN ARE PRESENT
                loopCount = str(i)
                f = open(logName, 'a') # Open Log File
                f.write('\n' + value + ',' + ' Account and Loss Run Present') #Log Success in File
                f.close() # Save and Close Log File
                print(loopCount + ' ' + value + ' Account and Loss Run Present')
                fnrenewals.append(value)
                accntstatus.append('Account and Loss Run Present')
                i = i + 1 # Increment i to count and track loops
                time.sleep(sleepTimer)
                chrome.quit()
##          ACCOUNT AND LOSS RUN BOTH EXIST
            except:
#WHEN ACCOUNT IS PRESENT BUT NOT LOSS RUN
                loopCount = str(i)
                f = open(logName, 'a') # Open Log File
                f.write('\n' + value + ',' + ' Account only') #Log Success in File
                f.close() # Save and Close Log File
                print(loopCount + ' ' + value, ' Account only')
                fnrenewals.append(value)
                accntstatus.append('Account only')
                i = i + 1 # Increment i to count and track loops
                time.sleep(sleepTimer)
                chrome.quit()
##          ONLY THE ACCOUNT EXISTS
        except:
#WHEN NEITHER ACCOUNT OR LOSS RUN ARE PRESENT
            loopCount = str(i)
            f = open(logName, 'a') # Open Log File
            f.write('\n' + value + ',' + ' No account or loss run') #Log Success in File
            f.close() # Save and Close Log File
            print(loopCount + ' ' + value, ' No account or loss run')
            fnrenewals.append(value)
            accntstatus.append('No account or loss run')
            i = i + 1 # Increment i to count and track loops
            time.sleep(sleepTimer)
            chrome.quit()
##      NO ACCOUNT OR LOSS RUN EXISTS
    except Exception as errException:
        loopCount = str(i) # Get Loops and Convert to string
        timeStamp = str(time.ctime()) # Get Current Time and Convert to String
        errMsg = str(errException) # Get Error Message and Convert into String
        f = open(logName, 'a') # Open Log File
        f.write('\n' + value + ' Errored! -' + errMsg + timeStamp) #Log Error in File
        f.close() # Save and Close Log File
        print(errException)
        print(i) # Print Loop Number 
        i = i + 1 # Increment i
        time.sleep(5)
        chrome.quit() # End Sessions to ensure loop resets fully.
        
#print(fnrenewals)
#print(accntstatus)
status = dict(zip(fnrenewals, accntstatus))
#print(status)
dfem['Account Status'] = dfem['Mnfn Code'].map(status)

# EXPORT REFORMATTED ACCOUNT RENEWAL LIST TO EXCEL - this will be used to update all accounts with new assumed policies for A&H and Marine programs
dfem.to_excel(r'C:\Users\aggood1\Desktop\Python experiments\MAX\MAX Renewal List.xlsx', sheet_name='MAX Renewal List', index = False) # this also shows which accounts were NOT present on webapp before the script sets them up

#BEGIN SETTING UP NEW ACCOUNTS ON WEBAPP
dfem.drop(dfem[(dfem['Account Status'] == 'Account and Loss Run Present') | (dfem['Account Status'] == 'Account only')].index, inplace=True)
#print(dfem)

# ACCOUNT SETUP ARRAYS
mnfnNumbers = dfem['Mnfn Code'].tolist()
#print(mnfnNumbers)
accountNames = dfem['Customer Name'].tolist()
policyInceptions = dfem['Effective Date'].tolist()
renewalDates = dfem['Renewal Date'].tolist()
programTypes = dfem['Line Of Business'].tolist() #Coverage
accntOwner = 'good'


for accountName, mnfnNumber, programType, policyInception, renewalDate in zip(accountNames, mnfnNumbers, programTypes, policyInceptions, renewalDates) :
    try:
        #OPEN INTERNET EXPLORER
        chrome = webdriver.Ie() #changed to internet explorer because Chrome does not have all elements showing
        actions = ActionChains(chrome)
        wait = WebDriverWait(chrome, 10)

        #GO TO MLNR WEB PAGE
        chrome.get('http://webapplink.com/Default.aspx') #goes to the home page
        time.sleep(sleepTimer)
        chrome.get('http://webapplink.com/AccountRef.aspx') #goes to the account reference page

        #PRODUCTION ENVIRONMENT
        #chrome.get('http://webapplink.com/Default.aspx') #goes to the home page
        #time.sleep(sleepTimer)
        #chrome.get('http://webapplink.com/AccountRef.aspx') #goes to the account reference page


        #CLICK ADD ACCOUNT
        element = wait.until(EC.element_to_be_clickable((By.NAME, 'add$accnt$btn')))
        chrome.find_element_by_name('add$accnt$btn').click() #presses the add account number
        element = wait.until(EC.element_to_be_clickable((By.NAME, 'accnt$name$text')))
        time.sleep(sleepTimer)

        #ENTER THE ACCOUNT DETAILS IN ACCOUNT PAGE - Setup Account
        chrome.find_element_by_name('accnt$name$text').send_keys(accountName) #ENTERS THE ACCOUNT NAME
        chrome.find_element_by_name('policy$type').send_keys(programType) #Line of Business
        chrome.find_element_by_name('prod$ent').send_keys("AOG")
        chrome.find_element_by_name('Accnt$Ownr').send_keys(accntOwner) #analyst
        chrome.find_element_by_name('Rvw$Date').send_keys(datetime.datetime.today().strftime('%m/%d/%Y'))
        chrome.find_element_by_name('Save$Btn').send_keys(Keys.RETURN)
        time.sleep(sleepTimer)

        #GET THE ACCOUNT NUMBER YOU'VE JUST CREATED AND STORE IT INTO A VARIABLE
        account_id = chrome.find_element_by_name('accnt$number$text').get_attribute('value')
        print(account_id)
        
        #POPULATES JAVASCRIPT TABS IN THE ACCOUNT PAGE - Add MNCode
        element = wait.until(EC.element_to_be_clickable((By.ID, 'AddMNCode')))
        time.sleep(sleepTimer)
        chrome.find_element_by_id('AddMNCode').click()
        element = wait.until(EC.element_to_be_clickable((By.NAME, 'MN$Code$text')))
        chrome.find_element_by_name('MN$Code$text').send_keys(mnfnNumber)
        chrome.find_element_by_name('Save$MN$Code').click()
        time.sleep(sleepTimer)
        
        #NAVIGATE TO LOSS RUNS TAB
        element = wait.until(EC.element_to_be_clickable((By.ID, 'TabOne')))
        time.sleep(sleepTimer)
        chrome.find_element_by_id("TabOne").click()

        #CHECK IF LOSS RUN EXISTS
        time.sleep(sleepTimer)
        isPresent = 0 # Reset Counter
        isPresent = chrome.find_elements_by_id("LossIDbutton") #Check for Loss Run and push into Array

        #IF LOSSRUN EXISTS, STOP SETUP AND MOVE ON TO NEXT ONE
        if len(isPresent) > 0 : #Check Array Length to determine if Loss Run Exists Already
            print('FIRST ERROR EXCEPTION CAUGHT')
            timeStamp = str(time.ctime()) # Get Current Time
            f = open(logName, 'a') # Open Log File
            f.write('\n' + accountName + ' ' + mnfnNumber + ' Errored! - Loss Run Already Exists ' + timeStamp) #Log Loss Run Error in File
            f.close() # Save and Close Log File
            print('Loss Run already Exists!')
            print(i) # Print Loop Number 
            i = i + 1 # Increment i
            time.sleep(sleepTimer+1)
            chrome.quit()
            continue

        #ELSE CREATE LOSS RUN
        else:
            #CREATE NEW LOSS RUN
            time.sleep(sleepTimer)
            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Create$LossRn$btn')))
            chrome.find_element_by_name("Create$LossRn$btn").click()

            #LOCATE ALERT BOX
            time.sleep(sleepTimer)
            element = wait.until(EC.alert_is_present())
            alert = chrome.switch_to_alert()
            alert.accept()
            time.sleep(sleepTimer)
            element = wait.until(EC.alert_is_present())
            alert = chrome.switch_to_alert()
            alert.accept()

            #OPEN CREATED LOSS RUN
            time.sleep(sleepTimer)
            element = wait.until(EC.element_to_be_clickable((By.ID, 'LossIDbutton')))
            target = chrome.find_element_by_id("LossIDbutton")
            str1="http://webapplink.com/AccountDefinition.aspx?AccntId="
            str2=target.text
            chrome.get(str1+str2)

            #EDIT FIRST POST
            time.sleep(sleepTimer)
            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Post$Edit$btn')))
            chrome.find_element_by_name("Post$Edit$btn").click()
            time.sleep(sleepTimer)
            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Comment$text')))
            chrome.find_element_by_name("Comment$text").send_keys('FN View 1 Direct Chubb Affiliates') #Comment box
            chrome.find_element_by_name("Mnfn$list$options").send_keys('Option1') #MN/FN CODE VIEW
            if (programType == "A&H" or programType == "Marine"):
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$list$options')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$list$options')))
                    chrome.find_element_by_name("Coverage$list$options").send_keys('include') #Coverage
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$list$options')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$list$options')))
                    chrome.find_element_by_name("Coverage$list$options").send_keys('include')  # Coverage
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$btn')))
                    chrome.find_element_by_name("Coverage$btn").click() #Coverage
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$btn')))
                    chrome.find_element_by_name("Coverage$btn").click()  # Coverage
                #LINE OF BUSINESS SELECTION
                if programType == 'A&H':
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'AH$chkbox')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'AH$chkbox')))
                        chrome.find_element_by_name("AH$chkbox").click()  # click on A&H
                        time.sleep(1)
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'AH$chkbox')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'AH$chkbox')))
                        chrome.find_element_by_name("AH$chkbox").click()  # click on A&H
                        time.sleep(1)
                elif programType == 'Marine':
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Marine$chkbox')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Marine$chkbox')))
                        chrome.find_element_by_name("Marine$chkbox").click()  # click on A&H
                        time.sleep(1)
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Marine$chkbox')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Marine$chkbox')))
                        chrome.find_element_by_name("Marine$chkbox").click()  # click on A&H
                        time.sleep(1)
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Inland$chkbox')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Inland$chkbox')))
                        chrome.find_element_by_name("Inland$chkbox").click()
                        time.sleep(1)
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Inland$chkbox')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Inland$chkbox')))
                        chrome.find_element_by_name("Inland$chkbox").click()
                        time.sleep(1)
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'OK$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'OK$btn')))
                    chrome.find_element_by_name("OK$btn").click() # click on update
                    time.sleep(1)
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'OK$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'OK$btn')))
                    chrome.find_element_by_name("OK$btn").click()  # click on update
                    time.sleep(1)
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Save$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                    chrome.find_element_by_name("Save$btn").click() # click on save LOB choices
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Save$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                    chrome.find_element_by_name("Save$btn").click()  # click on save LOB choices
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Cancel$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Cancel$btn')))
                    chrome.find_element_by_name("Cancel$btn").click() # click on cancel button to exit post criteria
                    time.sleep(1)
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Cancel$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Cancel$btn')))
                    chrome.find_element_by_name("Cancel$btn").click()  # click on cancel button to exit post criteria
                    time.sleep(1)
            ##INDENT FOR ALL COVERAGES EXCEPT FOR A&H AND MARINE
            else:
                if (programType == "Financial Lines / Professional Risk" or programType == "Environmental Risk" or programType == "Multiple Lines of Business"):
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$list$options')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$list$options')))
                        chrome.find_element_by_name("Coverage$list$options").send_keys('all') #Coverage:
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$list$options')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$list$options')))
                        chrome.find_element_by_name("Coverage$list$options").send_keys('all')  # Coverage:
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Save$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                        chrome.find_element_by_name("Save$btn").click() # click on save LOB choices
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Save$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                        chrome.find_element_by_name("Save$btn").click()  # click on save LOB choices
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Cancel$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Cancel$btn')))
                        chrome.find_element_by_name("Cancel$btn").click() # click on cancel button to exit post criteria
                        time.sleep(1)
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Cancel$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Cancel$btn')))
                        chrome.find_element_by_name("Cancel$btn").click()  # click on cancel button to exit post criteria
                        time.sleep(1)
                else:
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$list$options')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$list$options')))
                        chrome.find_element_by_name("Coverage$list$options").send_keys('include') #Coverage
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$list$options')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$list$options')))
                        chrome.find_element_by_name("Coverage$list$options").send_keys('include')  # Coverage
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$btn')))
                        chrome.find_element_by_name("Coverage$btn").click() #Coverage
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'Coverage$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'Coverage$btn')))
                        chrome.find_element_by_name("Coverage$btn").click()  # Coverage
                    

                #LINE OF BUSINESS SELECTION        
                    if programType == 'Property':
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Property$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Property$chkbox')))
                            chrome.find_element_by_name("Property$chkbox").click()
                            time.sleep(1)# click on Property
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Property$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Property$chkbox')))
                            chrome.find_element_by_name("Property$chkbox").click()
                            time.sleep(1)
                    elif programType == 'Casualty':
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Casualty$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Casualty$chkbox')))
                            chrome.find_element_by_name("Casualty$chkbox").click()
                            time.sleep(1)
                            # click on Casualty
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Casualty$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Casualty$chkbox')))
                            chrome.find_element_by_name("Casualty$chkbox").click()
                            time.sleep(1)
                            # click on Casualty
                    elif programType == 'A&H':
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'AH$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'AH$chkbox')))
                            chrome.find_element_by_name("AH$chkbox").click()  # click on A&H
                            time.sleep(1)
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'AH$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'AH$chkbox')))
                            chrome.find_element_by_name("AH$chkbox").click()  # click on A&H
                            time.sleep(1)
                    elif programType == 'Marine':
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Marine$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Marine$chkbox')))
                            chrome.find_element_by_name("Marine$chkbox").click()  # click on A&H
                            time.sleep(1)
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Marine$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Marine$chkbox')))
                            chrome.find_element_by_name("Marine$chkbox").click()  # click on A&H
                            time.sleep(1)
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Inland$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Inland$chkbox')))
                            chrome.find_element_by_name("Inland$chkbox").click()
                            time.sleep(1)
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Inland$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Inland$chkbox')))
                            chrome.find_element_by_name("Inland$chkbox").click()
                            time.sleep(1)
                    elif programType == 'Casualty / Property':
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Property$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Property$chkbox')))
                            chrome.find_element_by_name("Property$chkbox").click()
                            time.sleep(1)# click on Property
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Property$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Property$chkbox')))
                            chrome.find_element_by_name("Property$chkbox").click()
                            time.sleep(1)
                        try:
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Casualty$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Casualty$chkbox')))
                            chrome.find_element_by_name("Casualty$chkbox").click()
                            time.sleep(1)
                            # click on Casualty
                        except:
                            for handle in chrome.window_handles:
                                chrome.switch_to_window(handle)
                            element = wait.until(EC.presence_of_element_located((By.NAME, 'Casualty$chkbox')))
                            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Casualty$chkbox')))
                            chrome.find_element_by_name("Casualty$chkbox").click()
                            time.sleep(1)
                            # click on Casualty
                    try:
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'OK$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'OK$btn')))
                        chrome.find_element_by_name("OK$btn").click() # click on update
                        time.sleep(1)
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'OK$btn')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'OK$btn')))
                        chrome.find_element_by_name("OK$btn").click()  # click on update
                        time.sleep(1)
                

                    #CLOSE THE POST
                    time.sleep(sleepTimer)
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                    chrome.find_element_by_name("Save$btn").click() # click on cancel button to exit post criteria
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Cancel$btn')))
                    chrome.find_element_by_name("Cancel$btn").click() # click on cancel button to exit post criteria
                    time.sleep(sleepTimer)

                #DUPLICATE FIRST POST
                
                try:
                    element = wait.until(EC.presence_of_element_located((By.ID, 'PostEditbtn')))
                    element = wait.until(EC.element_to_be_clickable((By.ID, 'PostEditbtn')))
                    chrome.find_element_by_id("PostEditbtn").click() #GETS INFO ON FIRST POST
                    time.sleep(1)
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Post$Edit$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Post$Edit$btn')))
                    chrome.find_element_by_name("Post$Edit$btn").click()  # GETS INFO ON FIRST POST
                    time.sleep(1)
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Save$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                    chrome.find_element_by_name("Save$btn").click() #hits the duplicate button
                    time.sleep(1)
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Save$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                    chrome.find_element_by_name("Save$btn").click()  # hits the duplicate button
                    time.sleep(1)
                #START POPULATING FORM OF SECOND POST
                try:
                    chrome.find_element_by_name("Comment$text").clear()
                    time.sleep(1)
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    chrome.find_element_by_name("Comment$text").clear()
                    time.sleep(2)
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'Comment$text')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'Comment$text')))
                    chrome.find_element_by_name("Comment$text").send_keys('FN View 1 Assumed Excl Chubb Affiliates') #MN/FN CODE entry
                    chrome.find_element_by_name("ctl00$contentpalce$ddlDAC").send_keys('Assumed') #MN/FN CODE VIEW
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    chrome.find_element_by_name("Comment$text").clear()
                    chrome.find_element_by_name("ctl00$contentpalce$ddlDAC").clear()
                    chrome.find_element_by_name("Comment$text").send_keys('FN View 1 Assumed Excl Chubb Affiliates')  # MN/FN CODE entry
                    chrome.find_element_by_name("ctl00$contentpalce$ddlDAC").send_keys('Assumed')  # MN/FN CODE VIEW
                #CHOOSE THE COUNTRY OF LOSS
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'LossCountry$list$options')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'LossCountry$list$options')))
                    chrome.find_element_by_name("LossCountry$list$options").send_keys('exclude')
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'LossCountry$list$options')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'LossCountry$list$options')))
                    chrome.find_element_by_name("LossCountry$list$options").send_keys('exclude')
                try:
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'LossCountry$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'LossCountry$btn')))
                    chrome.find_element_by_name("LossCountry$btn").click() #Click country of loss box
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'LossCountry$btn')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'LossCountry$btn')))
                    chrome.find_element_by_name("LossCountry$btn").click()  # Click country of loss box
                try:
                    element = wait.until(EC.presence_of_element_located((By.ID, 'A')))
                    element = wait.until(EC.element_to_be_clickable((By.ID, 'A')))
                except:
                    for handle in chrome.window_handles:
                        chrome.switch_to_window(handle)
                    element = wait.until(EC.presence_of_element_located((By.ID, 'A')))
                    element = wait.until(EC.element_to_be_clickable((By.ID, 'A')))
                    #COUNTRY SELECTION
                #A
                chrome.find_element_by_id("LossCountry_001_chkbox").click() #Country 1
                chrome.find_element_by_id("LossCountry_002_chkbox").click() #Country 2
                chrome.find_element_by_id("LossCountry_003_chkbox").click() #Country 3
                #B
                chrome.find_element_by_id("B").click() #Click B
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_004_chkbox").click() #Country 4
                chrome.find_element_by_id("LossCountry_005_chkbox").click() #Country 5
                chrome.find_element_by_id("LossCountry_006_chkbox").click() #Country 6
                #C
                chrome.find_element_by_id("C").click() #Click C
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_007_chkbox").click() #Country 7
                chrome.find_element_by_id("LossCountry_008_chkbox").click() #Country 8
                #D
                chrome.find_element_by_id("D").click() #Click D
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_009_chkbox").click() #Country 9
                #E
                chrome.find_element_by_id("E").click() #Click E
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_010_chkbox").click() #Country 10
                chrome.find_element_by_id("LossCountry_011_chkbox").click() #Country 11
                #F
                chrome.find_element_by_id("F").click() #Click F
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_012_chkbox").click() #Country 12
                chrome.find_element_by_id("LossCountry_013_chkbox").click() #Country 13
                #G
                chrome.find_element_by_id("G").click() #Click G
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_014_chkbox").click() #Country 14
                chrome.find_element_by_id("LossCountry_015_chkbox").click() #Country 15
                chrome.find_element_by_id("LossCountry_016_chkbox").click() #Country 16
                #H
                chrome.find_element_by_id("H").click() #Click H
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_017_chkbox").click() #Country 17
                chrome.find_element_by_id("LossCountry_018_chkbox").click() #Country 18
                chrome.find_element_by_id("LossCountry_019_chkbox").click() #Country 19
                #I
                chrome.find_element_by_id("I").click() #Click I
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_020_chkbox").click() #Country 20
                chrome.find_element_by_id("LossCountry_021_chkbox").click() #Country 21
                chrome.find_element_by_id("LossCountry_022_chkbox").click() #Country 22
                chrome.find_element_by_id("LossCountry_023_chkbox").click() #Country 23
                chrome.find_element_by_id("LossCountry_024_chkbox").click() #Country 24
                chrome.find_element_by_id("LossCountry_025_chkbox").click() #Country 25
                #J
                chrome.find_element_by_id("J").click() #Click J
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_026_chkbox").click() #Country 26
                chrome.find_element_by_id("LossCountry_027_chkbox").click() #Country 27
                #K
                chrome.find_element_by_id("K").click() #Click K
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_028_chkbox").click() #Country 28
                #L
                chrome.find_element_by_id("L").click() #Click L
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_029_chkbox").click() #Country 29
                chrome.find_element_by_id("LossCountry_030_chkbox").click() #Country 30
                #M
                chrome.find_element_by_id("M").click() #Click M
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_031_chkbox").click() #Country 31
                chrome.find_element_by_id("LossCountry_032_chkbox").click() #Country 32
                #N
                chrome.find_element_by_id("N").click() #Click N
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_033_chkbox").click() #Country 33
                chrome.find_element_by_id("LossCountry_034_chkbox").click() #Country 34
                chrome.find_element_by_id("LossCountry_035_chkbox").click() #Country 35
                chrome.find_element_by_id("LossCountry_036_chkbox").click() #Country 36
                #P
                chrome.find_element_by_id("P").click() #Click P
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_037_chkbox").click() #Country 37
                #R
                chrome.find_element_by_id("R").click() #Click R
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_038_chkbox").click() #Country 38
                chrome.find_element_by_id("LossCountry_039_chkbox").click() #Country 39
                chrome.find_element_by_id("LossCountry_040_chkbox").click() #Country 40
                #S
                chrome.find_element_by_id("S").click() #Click S
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_041_chkbox").click() #Country 41
                chrome.find_element_by_id("LossCountry_042_chkbox").click() #Country 42
                chrome.find_element_by_id("LossCountry_043_chkbox").click() #Country 43
                chrome.find_element_by_id("LossCountry_044_chkbox").click() #Country 44
                chrome.find_element_by_id("LossCountry_045_chkbox").click() #Country 45
                chrome.find_element_by_id("LossCountry_046_chkbox").click() #Country 46
                chrome.find_element_by_id("LossCountry_047_chkbox").click() #Country 47
                chrome.find_element_by_id("LossCountry_048_chkbox").click() #Country 48
                chrome.find_element_by_id("LossCountry_049_chkbox").click() #Country 49
                chrome.find_element_by_id("LossCountry_050_chkbox").click() #Country 50
                #T
                chrome.find_element_by_id("T").click() #Click T
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_051_chkbox").click() #Country 51
                chrome.find_element_by_id("LossCountry_052_chkbox").click() #Country 52
                #U
                chrome.find_element_by_id("U").click() #Click U
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_053_chkbox").click() #Country 53
                chrome.find_element_by_id("LossCountry_054_chkbox").click() #Country 54
                chrome.find_element_by_id("LossCountry_055_chkbox").click() #Country 55
                chrome.find_element_by_id("LossCountry_056_chkbox").click() #Country 56
                #V
                chrome.find_element_by_id("V").click() #Click V
                time.sleep(1)
                chrome.find_element_by_id("LossCountry_057_chkbox").click() #Country 57

                chrome.find_element_by_name("Update$btn").click() # click on save LOB choices
                time.sleep(1)

                #CLOSE THE POST
                time.sleep(sleepTimer)
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'Save$btn')))
                chrome.find_element_by_name("Save$btn").click() # click on cancel button to exit post criteria
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'Cancel$btn')))
                chrome.find_element_by_name("Cancel$btn").click() # click on cancel button to exit post criteria
                time.sleep(sleepTimer)

            ##END A&H AND MARINE INDENT   

            #EDIT LOSS RUN TITLE/DESCRIPTION
            time.sleep(sleepTimer)
            element = wait.until(EC.element_to_be_clickable((By.NAME, 'Accnt$Def$text')))
            chrome.find_element_by_name("Accnt$Def$text").clear() 
            chrome.find_element_by_name("Accnt$Def$text").send_keys('Int\'l ', programType, ' Loss Information Since ', policyInception)

            #Inception date
            try:
                element = wait.until(EC.presence_of_element_located((By.NAME, 'Program$Edit$btn')))
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'Program$Edit$btn')))
                chrome.find_element_by_name("Program$Edit$btn").click()
            except:
                for handle in chrome.window_handles:
                    chrome.switch_to_window(handle)
                chrome.find_element_by_name("Program$Edit$btn").click()
            try:
                element = wait.until(EC.presence_of_element_located((By.NAME, 'Program$text')))
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'Program$text')))
                chrome.find_element_by_name("Program$text").clear()
            except:
                for handle in chrome.window_handles:
                    chrome.switch_to_window(handle)
                chrome.find_element_by_name("Program$text").clear()
            try:
                element = wait.until(EC.presence_of_element_located((By.NAME, 'Program$text')))
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'Program$text')))
                chrome.find_element_by_name("Program$text").send_keys(renewalDate)
            except:
                for handle in chrome.window_handles:
                    chrome.switch_to_window(handle)
                chrome.find_element_by_name("Program$text").send_keys(renewalDate)
            try:
                element = wait.until(EC.presence_of_element_located((By.NAME, 'Program$Prd$btn')))
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'Program$Prd$btn')))
                chrome.find_element_by_name("Program$Prd$btn").click()
                time.sleep(1)
            except:
                for handle in chrome.window_handles:
                    chrome.switch_to_window(handle)
                chrome.find_element_by_name("Program$Prd$btn").click()
                time.sleep(1)

            #SAVE LOSS RUN
            chrome.find_element_by_name("Save$btn").click()

            #CLOSE ALL SESSIONS
            loopCount = str(i)
            f = open(logName, 'a') # Open Log File
            f.write('\n' + loopCount + ' ' + accountName + ' ' + mnfnNumber + ' Setup Successfully') #Log Success in File
            f.close() # Save and Close Log File
            print (loopCount + accountName + ' Setup Successfully')
            i = i + 1 # Increment i to count and track loops
            time.sleep(sleepTimer+1)
            chrome.quit() # End Sessions
            

        # CATCH ALL EXCEPTIONS AND LOG IT
    except Exception as errException:
        print('SECOND ERROR EXCEPTION CAUGHT')
        time.sleep(sleepTimer)
        loopCount = str(i) # Get Loops and Convert to string
        timeStamp = str(time.ctime()) # Get Current Time and Convert to String
        errMsg = str(errException) # Get Error Message and Convert into String
        f = open(logName, 'a') # Open Log File
        f.write('\n' + loopCount + ' ' + accountName + ' ' + mnfnNumber + ' Errored! -' + errMsg + timeStamp) #Log Error in File
        f.close() # Save and Close Log File
        print(errException)
        chrome.get('http://webapplink.com/AccountRef.aspx') #goes to the loss run page
        wait = WebDriverWait(chrome, 5)
        chrome.find_element_by_name('search$account').send_keys(account_id) #ENTERS THE ACCOUNT NUMBER
        chrome.find_element_by_name('search$button').click() #presses the search button
        try:
            try:
                element = wait.until(EC.presence_of_element_located((By.NAME, 'account$edit$button')))
                element = wait.until(EC.element_to_be_clickable((By.NAME, 'account$edit$button')))
                time.sleep(sleepTimer)
#FIND THE LOSS RUN
                chrome.find_element_by_name('account$edit$button').click() #edit account
                time.sleep(sleepTimer)
                try:
                    try:
                        element = wait.until(EC.presence_of_element_located((By.ID, 'TabOne')))
                        element = wait.until(EC.element_to_be_clickable((By.ID, 'TabOne')))
                        chrome.find_element_by_id("TabOne").click()
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.ID, 'TabOne')))
                        element = wait.until(EC.element_to_be_clickable((By.ID, 'TabOne')))
                        chrome.find_element_by_id("TabOne").click()
                    try:
                        element = wait.until(EC.presence_of_element_located((By.ID, 'LossIDbutton')))
                        element = wait.until(EC.element_to_be_clickable((By.ID, 'LossIDbutton')))
                        target = chrome.find_element_by_id("LossIDbutton")
                        LossID = target.text
                    except:
                        for handle in chrome.window_handles:
                            chrome.switch_to_window(handle)
                        element = wait.until(EC.presence_of_element_located((By.NAME, 'LossIDbutton')))
                        element = wait.until(EC.element_to_be_clickable((By.NAME, 'LossIDbutton')))
                        target = chrome.find_element_by_id("LossIDbutton")
                        LossID = target.text
#WHEN BOTH ACCOUNT AND LOSS RUN ARE PRESENT
                    f = open(logName, 'a') # Open Log File
                    f.write('\n' + account_id + LossID + ', Account and Loss Run Present') #Log Success in File
                    f.close() # Save and Close Log File
                    print(account_id + ', ' + LossID + ', Account and Loss Run Present')
                    loglist.append(account_id + ', ' + LossID + ', Account and Loss Run Present')
                    time.sleep(sleepTimer)
                    #Enter the loss run
                    link="http://webapplink.com/AccountDefinition.aspx?AccntId="
                    chrome.get(link+LossID)
                    print("Entered Loss Run")
                    #Delete the loss run
                    time.sleep(sleepTimer)
                    chrome.find_element_by_id("Delete$btn").click() #hit the delete button
                    time.sleep(sleepTimer)
                    element = wait.until(EC.alert_is_present()) #Wait for alert to appear
                    alert = chrome.switch_to_alert() #Switch to the alert pop up
                    alert.accept() #Accept the alert pop up
                    time.sleep(5)
                    #Go back to the account page
                    chrome.get('http://webapplink.com/AccountRef.aspx')
                    wait = WebDriverWait(chrome, 5)
                    chrome.find_element_by_name('search$account').send_keys(account_id) #ENTERS THE ACCOUNT NUMBER
                    chrome.find_element_by_name('search$button').click() #presses the search button
                    element = wait.until(EC.presence_of_element_located((By.NAME, 'account$edit$button')))
                    element = wait.until(EC.element_to_be_clickable((By.NAME, 'account$edit$button')))
                    time.sleep(sleepTimer)
                    chrome.find_element_by_name('account$edit$button').click() #edit account
                    time.sleep(sleepTimer)
                    chrome.find_element_by_id("Delete$btn").click() #hit the delete button
                    time.sleep(sleepTimer)
                    element = wait.until(EC.alert_is_present()) #Wait for alert to appear
                    alert = chrome.switch_to_alert() #Switch to the alert pop up
                    alert.accept() #Accept the alert pop up
                    time.sleep(5)
                    loglist.append(account_id + ',' + ' Account and Loss Run Deleted')
                    print ('Account and Loss Run Deleted')
                    account_id = "Dummy Account"
                    LossID = "Dummy Loss"
                    chrome.quit()
##          ACCOUNT AND LOSS RUN BOTH EXIST
                except:
                    f = open(logName, 'a') # Open Log File
                    f.write('\n' + account_id + ', Account only') #Log Success in File
                    f.close() # Save and Close Log File
                    print(account_id, ' Account only')
                    loglist.append(account_id + ', Account only')
                    time.sleep(sleepTimer)
                    chrome.find_element_by_id("Delete$btn").click() #hit the delete button
                    time.sleep(sleepTimer)
                    element = wait.until(EC.alert_is_present()) #Wait for alert to appear
                    alert = chrome.switch_to_alert() #Switch to the alert pop up
                    alert.accept() #Accept the alert pop up
                    time.sleep(5)
                    loglist.append(account_id + ', Account Deleted')
                    print ('Account Deleted')
                    account_id = "Dummy Account"
                    LossID = "Dummy Loss"
                    chrome.quit()
##          ONLY THE ACCOUNT EXISTS
            except:
                
#WHEN NEITHER ACCOUNT OR LOSS RUN ARE PRESENT
                f = open(logName, 'a') # Open Log File
                f.write('\n' + account_id + ',' + ' Account Number does not exist') #Log Success in File
                f.close() # Save and Close Log File
                print(account_id, ' Account Number does not exist')
                loglist.append(account_id + ', Account Number does not exist')
                account_id = "Dummy Account"
                LossID = "Dummy Loss"
                time.sleep(sleepTimer)
                chrome.quit()
                #COPY UP TO HERE
        except Exception as errException:
            timeStamp = str(time.ctime()) # Get Current Time and Convert to String
            errMsg = str(errException) # Get Error Message and Convert into String
            f = open(logName, 'a') # Open Log File
            f.write('\n' + account_id + ' Errored! -' + errMsg + timeStamp) #Log Error in File
            f.close() # Save and Close Log File
            print(errException)
            time.sleep(5)
            chrome.quit() # End Sessions to ensure loop resets fully.
    
    continue
# END LOOP -----------------------------

# PRINT AND LOG COMPLETION NOTES -------
i = i - 1 #Because the loop always increments by 1, Last loop will increase increment by 1 more than there are actual items
loopCount = str(i)
timeStamp = str(time.ctime()) # Get Current Time
f = open(logName, 'a') # Open Log File
f.write('\n' + '0----------Script finished! at ' + timeStamp + '. Python tried to setup ' + loopCount + ' accounts') #Log Completion in File
f.close() # Save and Close Log File
print('0----------Script finished! at ' + timeStamp + '. Python tried to setup ' + loopCount + ' accounts. Please see '+ logName +' for any errors!')

# END OF SCRIPT