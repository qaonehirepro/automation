from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime
import time


class AssessmentUICommon:
    def __init__(self):
        pass

    def initiate_browser(self, url, path):
        self.driver = webdriver.Chrome(executable_path=path)
        self.driver.get(url)
        self.driver.implicitly_wait(10)
        self.driver.maximize_window()
        self.driver.switch_to.window(self.driver.window_handles[1])
        return self.driver

    def ui_login_to_test(self, user_name, password):

        time.sleep(10)
        self.driver.find_element(By.NAME, 'loginUsername').clear()
        self.driver.find_element(By.NAME, 'loginUsername').send_keys(user_name)
        self.driver.find_element(By.NAME, 'loginPassword').clear()
        self.driver.find_element(By.NAME, 'loginPassword').send_keys(password)
        self.driver.find_element(By.NAME, 'btnLogin').click()
        login_status = "None"
        try:
            if self.driver.find_element(By.XPATH,
                                        '//div[@class="text-center login-error ng-binding ng-scope"]').is_displayed():
                print("Unable to Login ")
                error_message = self.driver.find_element(By.XPATH,
                                                         '//div[@class="text-center login-error ng-binding ng-scope"]').text
                login_status = error_message
        except Exception as e:
            print(e)
            login_status = 'SUCCESS'
        return login_status

    def select_i_agree(self):
        time.sleep(1)
        i_agree_status = self.driver.find_element(By.CLASS_NAME, 'chk')
        is_selected = i_agree_status.is_selected()
        if not is_selected:
            i_agree_status.click()
            is_selected = True
        return is_selected

    def select_answer_for_the_question(self, answer):
        value = "//input[@name='answerOptions' and @value='%s']" % answer
        answered = self.driver.find_element(By.XPATH, value)
        is_answered = answered.is_selected()
        if not is_answered:
            answered.click()

    def check_answered_status(self, previous_answer):
        time.sleep(1)
        value = "//input[@name='answerOptions' and @value='%s']" % previous_answer
        answered = self.driver.find_element(By.XPATH, value).is_enabled()
        return answered

    def next_question(self, question_index):
        time.sleep(1)
        value = "btnQuestionIndex%s" % str(question_index)
        self.driver.find_element(By.NAME, value).click()

    def start_test_button_status(self):
        time.sleep(1)
        is_enabled = self.driver.find_element(By.NAME, 'btnStartTest').is_enabled()
        if is_enabled:
            start_button_status = 'Enabled'
        else:
            start_button_status = 'Disabled'
        return start_button_status

    def start_test(self):
        time.sleep(1)
        self.driver.find_element(By.NAME, 'btnStartTest').click()

    def end_test(self):
        time.sleep(3)
        self.driver.find_element(By.XPATH, "//button[@class='btn btn-danger ng-scope']").click()

    def end_test_confirmation(self):
        time.sleep(5)
        self.driver.find_element(By.NAME, 'btnCloseTest').click()
        print("Test is ended Successfully")

    def unanswer_question(self):
        self.driver.find_element(By.XPATH, "//button[@class='btn btn-default btnUnanswer ng-scope']").click()
        print("Un Answer Succeded")

    def find_question_string(self):
        question_string = self.driver.find_element(By.NAME, 'questionHtmlString').text
        print(question_string)
        return question_string


assess_ui_common_obj = AssessmentUICommon()
# status = assess_ui_common_obj.ui_login_to_test()
# print(status)
