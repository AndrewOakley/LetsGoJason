from selenium import webdriver
from selenium.webdriver.common.by import By
from xlwt import Workbook
from datetime import date, timedelta
from sys import exit
from sys import argv
from dotenv import load_dotenv
import os

load_dotenv(".env")

USERNAME = os.getenv("USER_NAME")
PASSWORD = os.getenv("PASS_WORD")

def login(driver):
    driver.find_element(by=By.XPATH, value='//*[@id="login-name"]').send_keys(USERNAME)
    driver.find_element(by=By.XPATH, value='//*[@id="password"]').send_keys(PASSWORD)
    driver.find_element(by=By.XPATH, value='//*[@id="login"]/div/form/div[3]/button').click()

    # remove this line if your name is Eric (line 18)
    # driver.find_element(by=By.XPATH, value='//*[@id="password"]/div[1]/a[2]').click()

def openJobs(driver, date=""):

    # check if there is a pending work order
    if(driver.find_element(by=By.XPATH, value='//*[@id="pending"]/div[2]/ul/li[2]').text[:19] == "Pending Work Orders"):
        # if there is a pending work order, click here for "Work Order Look up" button
        driver.find_element(by=By.XPATH, value='//*[@id="pending"]/div[2]/ul/li[9]/div/div/a').click()
    else:
        # otherwise click in the normal spot
        # click "Work Order Lookup" button
        driver.find_element(by=By.XPATH, value='//*[@id="pending"]/div[2]/ul/li[6]/div/div/a').click()

    # click "Find By Date" button
    driver.find_element(by=By.XPATH, value='//*[@id="lookup"]/div[2]/ul/li[3]/div/div/a').click()

    # check if the date is included
    # if it is, search for jobs by date, else just get today's jobs
    if date == "":
        # click "Find" button to get current days jobs
        driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[2]/form/div[2]/button').click()
    else:
        date = date.split('-')
        searchJobsByDate(driver, date)

def searchJobsByDate(driver, date):
    # click "Date Selection" button
    driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[2]/form/div[1]/div/a').click()
    # select year input and enter the given year
    yearInput = driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[5]/div[3]/input[1]')
    
    # select month input and enter the given month
    monthInput = driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[5]/div[3]/input[2]')
    # select day input and enter the given day
    dayInput = driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[5]/div[3]/input[3]')

    yearInput.click()
    yearInput.clear()
    yearInput.send_keys(date[0])
    monthInput.click()
    monthInput.clear()
    monthInput.send_keys(date[1].lstrip('0'))
    dayInput.click()
    dayInput.clear()
    dayInput.send_keys(date[2].lstrip('0'))

    # click "Set Date" button
    driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[5]/div[5]/a').click()
    # click "Find" button
    driver.find_element(by=By.XPATH, value='//*[@id="winforce-viewport"]/div[1]/div[2]/form/div[2]/button').click()

def getClientInfo(driver):
    clientPhoneNumber = ""
    clientEmail = ""
    clientInfo = []
    # check if there are notes on the job
    # this affects the layout of the page so different xpaths are used depending on the layout
    if driver.find_element(by=By.XPATH, value='//*[@id="view"]/div[2]/ul/li[2]').text == "WO Notes":
        # get client name and work number
        clientInfo = driver.find_element(by=By.XPATH, value=f'//*[@id="view"]/div[2]/ul/li[7]/div/div/a').text.split(' ')
        # click on user info button/tab
        driver.find_element(by=By.XPATH, value='//*[@id="view"]/div[2]/ul/li[7]/div/div/a').click()
    else:
        # get client name and work number
        clientInfo = driver.find_element(by=By.XPATH, value=f'//*[@id="view"]/div[2]/ul/li[5]/div/div/a').text.split(' ')
        # click a different xpath to get to the user info tab
        driver.find_element(by=By.XPATH, value='//*[@id="view"]/div[2]/ul/li[5]/div/div/a').click()
    
    # get client phone number
    clientPhoneNumber = ""
    try:
        clientPhoneNumber = driver.find_element(by=By.XPATH, value='//*[@id="subscriber"]/div[2]/ul/li[3]/div/div/a/nobr').text
    except:
        clientPhoneNumber = driver.find_element(by=By.XPATH, value='//*[@id="subscriber"]/div[2]/ul/li[4]/div/div/a/nobr').text
    # get client email
    clientEmail = ""
    try:
        clientEmail = driver.find_element(by=By.XPATH, value='//*[@id="subscriber"]/div[2]/ul/li[5]/div/div/a').text
    except:
        clientEmail = driver.find_element(by=By.XPATH, value='//*[@id="subscriber"]/div[2]/ul/li[4]/div/div/a').text
    # get the correct value for the client email
    clientEmail = clientEmail.split('\n')
    for i in clientEmail:
        if i.find("@") != -1:
            clientEmail = i
            break
    
    # click back button
    driver.find_element(by=By.XPATH, value='//*[@id="subscriber"]/div[1]/a[1]').click()

    return clientPhoneNumber, clientEmail, clientInfo

# log a single day of jobs
def logJobs(driver, excelSheet, row=0):
    noErrors = True
    index = 0
    row = row
    while noErrors:
        try:
            workNumber = driver.find_element(by=By.XPATH, value=f'//*[@id="winforce-viewport"]/div[1]/div[3]/ul/li[{index+2}]/div/div/a/nobr[1]').text
            # get formatted job date MM/DD
            currDate = driver.find_element(by=By.XPATH, value=f'//*[@id="winforce-viewport"]/div[1]/div[3]/ul/li[{index+2}]/div/div/a/nobr[2]').text
            jobDate = currDate[:-5].strip()
            
            # get job Address and city in correct format
            fullAddress = driver.find_element(by=By.XPATH, value=f'//*[@id="winforce-viewport"]/div[1]/div[3]/ul/li[{index+2}]/div/div/a/p').text.split(',')
            jobAddress = fullAddress[1] + "," + fullAddress[2]
            jobAddress.strip()
            jobCity = fullAddress[2][:-14].strip()

            # get client name and job Number
            # click on job to get more info
            driver.find_element(by=By.XPATH, value=f'//*[@id="winforce-viewport"]/div[1]/div[3]/ul/li[{index+2}]/div/div/a').click()
            clientPhoneNumber, clientEmail, clientInfo = getClientInfo(driver)
            clientNumber = clientInfo[0]
            clientName = (' '.join(clientInfo[1:])).upper()
            # click the back button to return to the list of jobs
            driver.find_element(by=By.XPATH, value=f'//*[@id="view"]/div[1]/a[1]').click()
            
            # write data to each column in excel sheet
            excelSheet.write(row, 0, workNumber)
            excelSheet.write(row, 1, jobDate)
            excelSheet.write(row, 2, jobAddress)
            excelSheet.write(row, 3, jobCity)
            excelSheet.write(row, 4, clientNumber)
            excelSheet.write(row, 5, clientName)
            excelSheet.write(row, 6, clientPhoneNumber)
            excelSheet.write(row, 7, clientEmail)

            print(row, workNumber, jobDate, jobAddress, jobCity, clientNumber, clientName, clientPhoneNumber, clientEmail)
            row +=1
            index += 1
        except Exception as e:
            print(F"THIS SCRIPT IS BAD AND MISSED SOME JOBS ON {currDate}")
            print("CHECK TO MAKE SURE THAT ALL JOBS ARE ACCOUNTED FOR (BECAUSE THEY PROBABLY AREN'T) !!!!!!")
            noErrors = False
            index += 1
    return row

# log jobs with a variable number of days
# if numDays is 7 it will log the last seven days not including today
def logJobsDelta(driver, excelSheet, currDate, numDays):
    row = 0
    date = currDate
    for i in range(numDays):
        row = logJobs(driver, excelSheet, row)
        # change date to next day
        date = date + timedelta(days=1)
        searchJobsByDate(driver, date.strftime('%Y-%m-%d').split("-"))

def displayHelp():
    print("\nEric McQuaid's Work Log Script Help:")
    print("\tE.G.\n\tpython logWork_v2.4.py --week")
    print()
    print("Running without arguments (python logWorkv2.4.py) will log todays jobs")
    print()
    print("Optional Arguments:")
    print("--help, -h\t\tDisplay this help message")
    print("--date=YYYY-MM-DD\tLog the given day's jobs")
    print("\t\t\tE.G.: python logWork_v2.4.py --date=2022-08-11")
    print("--week, -w\t\tLog the last 7 day's of jobs (not including today)")
    print("--month, -m\t\tLog the last 31 day's of jobs  (not including today)")
    print()
    pass

if __name__ == "__main__":
    currDate = date.today().strftime("%Y-%m-%d")
    # get command line arguments
    arguments = argv
    numDays = 0
    filename = ""
    # if there are more arguments than just the file name get the date
    if len(arguments) > 1:
        for args in arguments[1:]:
            if (args == "--help") or (args == "-h"):
                displayHelp()
                exit()
            elif args[:7] == "--date=":
                currDate = args[7:]
            elif (args == "--week") or (args == "-w"):
                print("Logging jobs for the Week...")
                numDays = 7
                filename = "wcmcquaid_weekof_"
            elif (args == "--month") or (args == "-m"):
                print("Logging jobs for the Month...")
                numDays = 31
                filename = "wcmcquaid_monthof_"
            else:
                print("\nInvalid arguments, follow this help guide to correctly run the program\n")
                displayHelp()
                exit()
    else:
        arguments = None

    # make True to clear datasets folder
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(options=options)
    

    driver.get("https://wft.glds.com/communicomm/users/login#login")
    login(driver)
    
    # log the full week or month of jobs on one sheet
    if numDays > 0:
        currDate = date.today() - timedelta(days=numDays)
        openJobs(driver, currDate.strftime("%Y-%m-%d"))
        filename = f"{filename}{date.today()}.xls"
        with open(filename, 'w') as f:
            # open excel sheet
            wb = Workbook()
            # add_sheet is used to create sheet.
            sheet1 = wb.add_sheet('Sheet 1', cell_overwrite_ok=True)
            logJobsDelta(driver, sheet1, currDate, numDays)

            wb.save(f.name)
    # log a single day of jobs on a single sheet
    else:
        openJobs(driver, currDate)
        filename = f"wcmcquaid_{currDate}.xls"
        with open(filename, 'w') as f:
            # open excel sheet
            wb = Workbook()
            # add_sheet is used to create sheet.
            sheet1 = wb.add_sheet('Sheet 1', cell_overwrite_ok=True)
            logJobs(driver, sheet1)

            wb.save(f.name)
    

    # close the browser window
    driver.close()
    print(f"Logged jobs on {currDate} to {filename}")
    exit()
