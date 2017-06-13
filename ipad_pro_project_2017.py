from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Ec
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException
from openpyxl import load_workbook
import time
import credentials

# Prompts user to select sheet within an Excel workbook
sheet_selection = input("Select sheet to work from: ")

# Lists that hold asset data from Excel sheet
name = []
corporate_id = []
asset = []
phone_number = []

username = credentials.login['username']
password = credentials.login['password']

# Points to Excel sheet which contains asset details based on user's previous selection
work_book = load_workbook('spreadsheets/iPad_Pro.xlsx')
work_book.get_sheet_names()
sheet = work_book.get_sheet_by_name(sheet_selection)

# Puts name, corporate id, asset, and phone number from predefined columns within Excel sheet
get_name = sheet['A']
for x in range(len(get_name)):
    name.append(str(get_name[x].value))

get_corporate_id = sheet['B']
for x in range(len(get_corporate_id)):
    corporate_id.append(str(get_corporate_id[x].value))

get_asset = sheet['C']
for x in range(len(get_asset)):
    asset.append(str(get_asset[x].value))

get_phone_number = sheet['D']
for x in range(len(get_phone_number)):
    number = str(get_phone_number[x].value)
    # Check for and corrects proper telephone number formatting
    if '-' not in number:
        number = number[:3] + "-" + number[3:6] + "-" + number[6:]
        phone_number.append(number)
    else:
        phone_number.append(number)

# Gets total number of entries made into the asset list
entry_count = len(get_asset)

# Sets up ChromeDriver and options
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.get("https://myit.ucb.com/ux/smart-it/#/")


# Performs the login into Smart IT
def smart_it_login(usn, pwd):
    u = WebDriverWait(driver, 60).until(Ec.presence_of_element_located((By.ID, 'loginUserName')))
    p = WebDriverWait(driver, 60).until(Ec.presence_of_element_located((By.ID, 'loginPass')))
    l = WebDriverWait(driver, 60).until(Ec.presence_of_element_located((By.XPATH, '//button[contains(text(), "Log In")]')))
    u.send_keys(usn)
    p.send_keys(pwd)
    l.click()


# Performs click action on specified html tags and re-attempts on fail
def execute_step(stp, cnt):
    try:
        stp = stp[1:]
        wait = WebDriverWait(driver, 60)
        element = wait.until(Ec.presence_of_element_located((By.XPATH, '%s' % stp)))
        element.click()
    except WebDriverException:
        print("Step %s has failed. 2 second delay set. Trying once more." % cnt)
        time.sleep(2)
        stp = "$" + stp
        execute_step(stp, cnt)


# Determines if step is an input action or click action
def by_xpath(step, count):
    steps_counter = count + 1
    print("%s | Step #%s" % (step, steps_counter))
    if step[:1] != "$":
        actions = ActionChains(driver)
        actions.send_keys(step)
        actions.send_keys(Keys.RETURN)
        actions.perform()

    else:
        execute_step(step, count)

smart_it_login(username, password)

# Counts iterations when used with ChromeDriver to close out browser window due to high memory usage buildup
mem_leak_counter = 0
for loop_count, a in enumerate(asset):
    if mem_leak_counter == 5:
        driver.quit()
        mem_leak_counter = 0
        driver = webdriver.Chrome(chrome_options=chrome_options)
        driver.get("https://myit.ucb.com/ux/smart-it/#/")
        smart_it_login(username, password)

    completed_count = loop_count + 1
    print("### %s out of %s entries completed ###" % (completed_count, entry_count))

    # List of all input and click steps required to complete a workflow
    steps = [
        '$//a[contains(text(), "Console")]',
        '$//span[contains(text(), "Asset Console")]',
        '$//button[contains(text(), "Clear Filters")]',
        '$//span[contains(text(), "Filter")]',
        '$//div[contains(text(), "Keywords")]',
        '$//input[contains(@placeholder, "Type a specific search term")]',
        asset[loop_count],
        '$//div[contains(text(), "Asset Type")]',
        '$//div[contains(text(), "Computer System")]',
        '$//div[contains(@ng-style, "rowStyle(row)")]',
        '$//div[contains(text(), "Status:")]',
        '$//div[contains(@title-text, "Asset Status")]',
        '$//a[contains(text(), "Deployed")]',
        '$//button[contains(text(), "Save")]',
        '$//a[contains(text(), "People")]',
        '$//div[contains(@ng-click, "addRelatedPeople()")]',
        '$//input[contains(@placeholder, "Type to search people")]',
        corporate_id[loop_count],
        '$//div[contains(@ng-click, "selectPerson(person)")]',
        '$//button[contains(text(), "Add People")]',
        '$//a[contains(text(), "Assets")]',
        '$//span[contains(text(), "Relate Existing Asset")]',
        '$//input[contains(@placeholder, "Search CIs by Name,ID,Serial Number")]',
        phone_number[loop_count],
        '$//input[contains(@type, "checkbox")]',
        '$//button[contains(@title, "Relationship Type Select one")]',
        '$//a[contains(text(), "Component")]',
        '$//button[contains(text(), "Save")]',
        '$//a[contains(text(), "Console")]',
        '$//span[contains(text(), "Asset Console")]',
        '$//button[contains(text(), "Clear Filters")]',
        '$//span[contains(text(), "Filter")]',
        '$//div[contains(text(), "Keywords")]',
        '$//input[contains(@placeholder, "Type a specific search term")]',
        phone_number[loop_count],
        '$//div[contains(text(), "Asset Type")]',
        '$//div[contains(text(), "Hardware")]',
        '$//div[contains(@ng-style, "rowStyle(row)")]',
        '$//div[contains(text(), "Status:")]',
        '$//div[contains(@title-text, "Asset Status")]',
        '$//a[contains(text(), "Deployed")]',
        '$//button[contains(text(), "Save")]',
        '$//a[contains(text(), "People")]',
        '$//div[contains(@ng-click, "addRelatedPeople()")]',
        '$//input[contains(@placeholder, "Type to search people")]',
        corporate_id[loop_count],
        '$//div[contains(@ng-click, "selectPerson(person)")]',
        '$//button[contains(text(), "Add People")]'
    ]

    for step_counter, s in enumerate(steps):
        by_xpath(s, step_counter)
    mem_leak_counter += 1
