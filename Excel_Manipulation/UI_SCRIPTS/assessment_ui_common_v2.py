from selenium import webdriver
import datetime
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException


class AssessmentUICommon:

    def __init__(self):
        self.delay = 120

    def initiate_browser(self, url, path):
        self.driver = webdriver.Chrome(executable_path=path)
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

    def select_i_agree(self):
        try:
            # time.sleep(1)
            i_agree_status = WebDriverWait(self.driver, self.delay).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'chk')))
            # i_agree_status = self.driver.find_element_by_class_name('chk')
            # i_agree_status = self.driver.find_element_by_xpath('//div[@class="chk"]')
            is_selected = i_agree_status.is_selected()
            if not is_selected:
                i_agree_status.click()
                is_selected = True
            return is_selected

        except Exception as e:
            print("I agree is not visible")
            print(e)

    def select_answer_for_the_question(self, answer):
        # time.sleep(1)
        value = "//input[@name='answerOptions' and @value='%s']" % answer
        answered = self.driver.find_element_by_xpath(value)
        is_answered = answered.is_selected()
        if not is_answered:
            answered.click()

    def check_answered_status(self, previous_answer):
        time.sleep(1)
        value = "//input[@name='answerOptions' and @value='%s']" % previous_answer
        answered = self.driver.find_element_by_xpath(value).is_enabled()
        return answered

    def next_question(self, question_index):
        time.sleep(1)
        value = "btnQuestionIndex%s" % str(question_index)
        self.driver.find_element_by_name(value).click()

    def start_test_button_status(self):
        time.sleep(1)
        is_enabled = self.driver.find_element_by_name('btnStartTest').is_enabled()
        if is_enabled:
            start_button_status = 'Enabled'
        else:
            start_button_status = 'Disabled'
        return start_button_status

    def start_test(self):
        time.sleep(1)
        self.driver.find_element_by_name('btnStartTest').click()

    def end_test(self):
        time.sleep(3)
        self.driver.find_element_by_xpath("//button[@class='btn btn-danger ng-scope']").click()

    def end_test_confirmation(self):
        time.sleep(5)
        self.driver.find_element_by_name('btnCloseTest').click()
        print("Test is ended Successfully")

    def unanswer_question(self):
        self.driver.find_element_by_xpath("//button[@class='btn btn-default btnUnanswer ng-scope']").click()
        print("Un Answer Succeded")

    def find_question_string(self):
        question_string = self.driver.find_element_by_name('questionHtmlString').text
        # print(question_string)
        groupname = self.driver.find_element_by_name('groupName').text
        section_name = self.driver.find_element_by_name('sectionName').text
        # print(groupname)
        # print(section_name)
        return question_string, groupname, section_name


assess_ui_common_obj = AssessmentUICommon()
# status = assess_ui_common_obj.ui_login_to_test()
# print(status)
