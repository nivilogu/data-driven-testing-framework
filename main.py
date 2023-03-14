# main.py

# This is the test script runable file

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from locators import Test_Locators
from excel_functions import Suman_Excel_Functions

excel_file = 'C:\\Users\\sgang\\OneDrive\\Desktop\\DDTF\\test_data.xlsx'

sheet_name = 'Sheet1'

s = Suman_Excel_Functions(excel_file, sheet_name)

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

driver.maximize_window()

driver.get('https://www.facebook.com')

rows = s.row_count()

for row in range(2, rows+1):
    username = s.read_data(row, 6)
    password = s.read_data(row, 7)

    driver.find_element(by=By.NAME, value=Test_Locators().username_locator).send_keys(username)
    driver.find_element(by=By.NAME, value=Test_Locators().password_locator).send_keys(password)
    driver.find_element(by=By.NAME, value=Test_Locators().submitButton_locator).click()
    
    driver.implicitly_wait(10)
    # write the test data into the excel file 
    if 'https://www.facebook.com/checkpoint/?next' in driver.current_url:
        print("SUCCESS : Login success with username {a}".format(a=username))
        s.write_data(row, 8, "TEST PASS")
        driver.back()
    elif('https://www.facebook.com' in driver.current_url):
        print("FAIL : Login FAILED with username {a}".format(a=username))
        s.write_data(row, 8, "TEST FAIL")
        driver.back()

driver.quit()


