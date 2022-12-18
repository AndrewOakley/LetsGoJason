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
    driver.find_element(by=By.CSS_SELECTOR, value="input[name='username']").click().send_keys("holland")
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/section/main/div/div/div[1]/div[2]/form/div/div[2]/div/label/input').click().send_keys("andy")
    driver.find_element(by=By.XPATH, value='//*[@id="loginForm"]/div/div[3]/button').click()

    # remove this line if your name is Eric (line 18)
    # driver.find_element(by=By.XPATH, value='//*[@id="password"]/div[1]/a[2]').click()

if __name__ == "__main__":

    # make True to clear datasets folder
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(options=options)

    driver.get("https://www.instagram.com/accounts/login/?next=%2Fgoogle%2F&source=desktop_nav&hl=en")
    login(driver)
    
    # close the browser window
    # driver.close()
    exit()
