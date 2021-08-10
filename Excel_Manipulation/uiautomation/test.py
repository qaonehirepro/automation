from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

browser = webdriver.Chrome('D:\\automation\\chromedriver.exe')
browser.get("https://amsin.hirepro.in/crpo/#/login/AT")
delay = 40 # seconds
try:
    myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.NAME, 'loginName')))
    print ("Page is ready!")
except TimeoutException:
    print("Hello")