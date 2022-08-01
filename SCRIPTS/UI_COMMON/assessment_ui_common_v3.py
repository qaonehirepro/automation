# from selenium import webdriver
# import datetime
import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# from selenium.common.exceptions import TimeoutException


class AssessmentUICommon:

    def __init__(self):
        self.delay = 120

    def initiate_browser(self, url, path):
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
        print(question_string)
        return question_string

    def rejection_page(self):
        data = {}
        try:
            if self.driver.find_element_by_name('nextTestMsg').is_displayed():
                message = self.driver.find_element_by_name('nextTestMsg').text
                overall_page_message = self.driver.find_element_by_xpath("//*[@class='ng-scope']").text
                data = {'is_next_test_available': 'Not Available', 'is_shortlisted': 'Rejected',
                        'message': message, 'consent_yes': 'EMPTY', 'consent_no': 'EMPTY',
                        'consent_paragraph': 'EMPTY', 'next_test_page_message': overall_page_message}

        except Exception as e:
            print(e)
            message = "shortlisting not available"
            data = {'is_next_test_available': 'EXCEPTION OCCURRED', 'is_shortlisted': 'EXCEPTION OCCURRED',
                    'message': message, 'consent_yes': 'EXCEPTION OCCURRED', 'consent_no': 'EXCEPTION OCCURRED',
                    'consent_paragraph': 'EXCEPTION OCCURRED', 'next_test_page_message': 'EXCEPTION OCCURRED'}
        return data

    def shortlisting_page(self):
        data = {}
        try:
            if self.driver.find_element_by_name('btnStartNextTest').is_displayed():
                overall_page_message = self.driver.find_element_by_xpath("//*[@class='ng-scope']").text
                button_message = self.driver.find_element_by_name('btnStartNextTest').text
                next_test_message = self.driver.find_element_by_name('nextTestMsg').text
                if button_message == 'Yes, Take me to the next test':
                    consent_message = self.driver.find_element_by_xpath("//*[@class='next-msg ng-scope']").text
                    consent_yes = self.driver.find_element_by_xpath("//*[@class='btn btn-success btn-yes']").text
                    consent_no = self.driver.find_element_by_xpath("//*[@class='btn btn-default red-button']").text
                    data = {'is_next_test_available': 'Available', 'is_shortlisted': 'Shortlisted with Consent',
                            'message': next_test_message, 'consent_yes': consent_yes, 'consent_no': consent_no,
                            'consent_paragraph': consent_message, 'next_test_page_message': overall_page_message}
                else:
                    if next_test_message == 'We have another test lined up for you.':
                        data = {'is_next_test_available': 'Available', 'is_shortlisted': 'Autotest',
                                'message': next_test_message, 'consent_yes': 'EMPTY', 'consent_no': 'EMPTY',
                                'consent_paragraph': 'EMPTY', 'next_test_page_message': overall_page_message}
                    elif next_test_message == 'Congratulations! You are eligible for the next test.':
                        data = {'is_next_test_available': 'Available', 'is_shortlisted': 'Shortlisted',
                                'message': next_test_message, 'consent_yes': 'EMPTY', 'consent_no': 'EMPTY',
                                'consent_paragraph': 'EMPTY', 'next_test_page_message': overall_page_message}
                    else:
                        data = {'is_next_test_available': 'Available', 'is_shortlisted': 'DEBUG',
                                'message': next_test_message, 'consent_yes': 'DEBUG', 'consent_no': 'DEBUG',
                                'consent_paragraph': 'DEBUG', 'next_test_page_message': overall_page_message}
        except Exception as e:
            print(e)
            next_test_message = "Shortlisting Not Available"
            data = {'is_next_test_available': 'EXCEPTION OCCURRED', 'is_shortlisted': 'EXCEPTION OCCURRED',
                    'message': next_test_message, 'consent_yes': 'EXCEPTION OCCURRED',
                    'consent_no': 'EXCEPTION OCCURRED', 'consent_paragraph': 'EXCEPTION OCCURRED',
                    'next_test_page_message': 'EXCEPTION OCCURRED'}
        print(data)
        return data

    def start_next_test(self):
        self.driver.find_element_by_name('btnStartNextTest').click()
        time.sleep(3)
        self.driver.switch_to.window(self.driver.window_handles[2])

    def consent_no(self):
        self.driver.find_element_by_xpath("//*[@class='btn btn-default red-button']").click()
        time.sleep(3)
        self.driver.switch_to.window(self.driver.window_handles[2])

    def vet_start_test(self):
        time.sleep(5)
        try:
            self.driver.switch_to_frame('thirdPartyIframe')
            self.driver.find_element_by_xpath("//*[@class='wdtContextualItem  wdtContextStart']").click()
            print("VET Test is started Successfully")
            vet_test_started = "Successful"
            is_element_successful = True

        except Exception as e:
            print(e)
            print("VET Start test is failed")
            is_element_successful = False
            vet_test_started = "Failed"
        return vet_test_started, is_element_successful

    def vet_welcome_page(self):
        time.sleep(5)
        try:
            self.driver.find_element_by_id('welcome_next_link').click()
            print("Welcome Page")
            vet_welcome_page = "Successful"
            is_element_successful = True
        except Exception as e:
            print(e)
            print("Failed in Welcome Page")
            is_element_successful = False
            vet_welcome_page = "Failed"

        return vet_welcome_page, is_element_successful

    def vet_quiet_please(self):
        time.sleep(5)
        try:
            self.driver.find_element_by_id('distraction_next_link').click()
            print("Quiet Please Page")
            vet_quiet_please_page = "Successful"
            is_element_successful = True

        except Exception as e:
            print(e)
            print("Welcome Page successful")
            vet_quiet_please_page = "Failed"
            is_element_successful = False

        return vet_quiet_please_page, is_element_successful

    def vet_ready_check_box(self):
        time.sleep(5)
        try:
            self.driver.find_element_by_id('ready_checkbox').click()
            print("Ready Check Box successfull")
            vet_ready_check_box = "Successful"
            is_element_successful = True
        except Exception as e:
            print(e)
            print("Ready Check Box Failed")
            vet_ready_check_box = "Failed"
            is_element_successful = False

        return vet_ready_check_box, is_element_successful

    def vet_ready_start_link(self):
        time.sleep(5)
        try:
            self.driver.find_element_by_id('ready_start_link').click()
            print("Ready Start Successful")
            vet_ready_check_box = "Successful"
            is_element_successful = True

        except Exception as e:
            print(e)
            print("vet_ready_start_link Failed")
            vet_ready_check_box = "Failed"
            is_element_successful = False
        return vet_ready_check_box, is_element_successful

    def vet_proceed_test(self):
        time.sleep(30)
        try:
            self.driver.find_element_by_xpath("//*[@class = 'proceed wizardButton greenBackground']").click()
            print("Proceed Test Successful")
            os.system("F:\\my_test.mp3")
            vet_proceed_test = "Successful"
            is_element_successful = True
            time.sleep(5)
        except Exception as e:
            print(e)
            print("vet_ready_start_link Failed")
            vet_proceed_test = "Failed"
            is_element_successful = False
        return vet_proceed_test, is_element_successful

    def vet_speaking_tips(self):
        try:
            time.sleep(120)
            self.driver.find_element_by_xpath("//*[@class='testInstructionItem testInstructionNext']").click()
            print("Speaking Tips Successful")
            vet_speaking_tips = "Successful"
            is_element_successful = False

        except Exception as e:
            print(e)
            print("vet_speaking_tips Failed")
            vet_speaking_tips = "Failed"
            is_element_successful = False

        return vet_speaking_tips, is_element_successful

    def vet_overview(self):
        try:
            time.sleep(10)
            self.driver.find_element_by_xpath("//*[@class='wdtContextualItem  wdtContextNext']").click()
            print("Overview Page Success")
            vet_overview = "Successful"
            is_element_successful = True

        except Exception as e:
            print(e)
            print("vet_speaking_tips Failed")
            vet_overview = "Failed"
            is_element_successful = False

        return vet_overview, is_element_successful

    def vet_instruction(self):
        time.sleep(10)
        try:
            self.driver.find_element_by_xpath("//*[@class='wdtContextualItem wdtContextNext']").click()
            print("VET Instructions Page Success")
            vet_instruction = "Successful"
            is_element_successful = True

        except Exception as e:
            print(e)
            print("VET Instructions Failed")
            vet_instruction = "Failed"
            is_element_successful = False

        return vet_instruction, is_element_successful

    def play_audio(self):
        os.system("F:\\my_test.mp3")

    def survey_submit(self):
        time.sleep(60)
        try:
            if self.driver.find_element_by_xpath("//*[@class = 'wdtContextualItem  wdtContextNext']").is_displayed():
                self.driver.find_element_by_xpath("//*[@class = 'wdtContextualItem  wdtContextNext']").click()
                print("survey Question Success")
            survey_submit = "Successful"
            is_element_successful = False
        except Exception as e:
            print(e)
            print("survey Question  Failed")
            survey_submit = "Failed"
            is_element_successful = False
        return survey_submit, is_element_successful


assess_ui_common_obj = AssessmentUICommon()
# status = assess_ui_common_obj.ui_login_to_test()
# print(status)
