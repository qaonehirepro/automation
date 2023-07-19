import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

class TenantLogin:
    def __init__(self):
        self.delay = 120

    def initiate_browser2(self, url, path):
        # chrome option is needed in VET cases - ( its handling permissions like mic access)
        chrome_options = Options()
        chrome_options.add_argument("--use-fake-ui-for-media-stream")
        self.driver = webdriver.Chrome(executable_path=path, chrome_options=chrome_options)
        self.driver.get(url)
        self.driver.implicitly_wait(10)
        self.driver.maximize_window()
        self.driver.switch_to.window(self.driver.window_handles[1])
        return self.driver

    def ui_login_to_test(self, user_name, password):
        time.sleep(5)
        self.driver.find_element_by_name('loginUsername').clear()
        self.driver.find_element_by_name('loginUsername').send_keys(user_name)
        self.driver.find_element_by_name('loginPassword').clear()
        self.driver.find_element_by_name('loginPassword').send_keys(password)
        self.driver.find_element_by_name('btnLogin').click()
        # time.sleep(5)
        login_status = "None"
        try:
            if self.driver.find_element_by_xpath(
                    '//div[@class="text-center login-error ng-binding ng-scope"]').is_displayed():
                print("Unable to Login ")
                error_message = self.driver.find_element_by_xpath(
                    '//div[@class="text-center login-error ng-binding ng-scope"]').text
                login_status = error_message
        except Exception as e:
            print(e)
            login_status = 'SUCCESS'
        return login_status
