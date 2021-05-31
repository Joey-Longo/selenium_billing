from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchWindowException, StaleElementReferenceException
import time
from pure_helper import logs, u_log


'''
Chrome driver 84 or greater required for headless download
otherwise comment out lines 16-17 and remove options from Chrome()
'''

chrome_path = '/PathTo/chromedriver'
chrome_options = Options()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)

driver.get('https://pure1.purestorage.com/')

# Pure FlashArrays
arrays = ['array1', 'array2']


def pure_login():
    login_page = driver.find_element_by_name('useremail')
    login_page.send_keys(u_log())

    driver.find_element_by_id('login-btn').click()
    try:
        login_page.find_element_by_class_name('form-control inputField pw')
    except StaleElementReferenceException:
        pass
    time.sleep(5)
    login_page_secondary = driver.find_element_by_name('loginPage:loginForm:j_id33')
    login_page_secondary.send_keys(logs())
    driver.find_element_by_id('loginPage:loginForm:doLogin').click()


def create_report():
    driver.get('https://pure1.purestorage.com/analysis/capacity/arrays')
    time.sleep(5)
    # Put array checkboxes into list
    box = driver.find_elements_by_class_name('value-cell-data')
    # Change data from percentage to bytes
    driver.find_element_by_id('arrayChartsUseAbsoluteRadioButton').click()
    # Go through all checkboxes and select only wanted arrays
    for i in range(len(box)):
        if box[i].text in arrays:
            driver.find_element_by_link_text(str(box[i].text)).click()
    # Window switch fails but still works anyway, added a pass to bypass any exceptions stopping the flow
    try:
        driver.switch_to.window('Pure1 Manage')
    except NoSuchWindowException:
        pass
    # Select 2 month time frame
    time_frame = driver.find_element_by_tag_name('select')
    choose = Select(time_frame)
    choose.select_by_value('1: Object')
    # Clicks export dropdown
    driver.find_element_by_class_name('export-text').click()
    # Put dropdown lists into a python list
    li = driver.find_elements_by_tag_name('li')
    # Run through li and click to export selected arrays
    for x in range(len(li)):
    # This work but throws a stale element exception, made a pass to prevent any stoppage
        try:
            if li[x].text == 'Export selected arrays (2)':
                li[x].click()
        except StaleElementReferenceException:
            pass
    # Final export confirm
    driver.find_element_by_id('exportButton').click()
    time.sleep(10)
    driver.quit()


pure_login()
time.sleep(15)
create_report()
