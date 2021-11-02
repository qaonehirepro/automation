# import datetime
#from COMMON.read_excel import *
from Excel_Manipulation.COMMON.read_excel import *
#from uiautomation.assessment_ui_common import *
from Excel_Manipulation.uiautomation.assessment_ui_common import *
import time


class OnlineAssessment:

    def __init__(self):
        self.url = "https://amsinsec.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0="
        self.path = r"F:\automation\chromedriver.exe"
        self.browser = assess_ui_common_obj.initiate_browser(self.url, self.path)

    def mcq_assessment(self, current_excel_data):
        login_details = assess_ui_common_obj.ui_login_to_test(current_excel_data.get('loginName'),
                                                              (current_excel_data.get('password')))
        if login_details == 'SUCCESS':
            i_agreed = assess_ui_common_obj.select_i_agree()
            if i_agreed:
                start_test_status = assess_ui_common_obj.start_test_button_status()
                assess_ui_common_obj.start_test()
                if current_excel_data.get('skipRquired') == 'Yes':
                    assess_ui_common_obj.next_question(2)
                    assess_ui_common_obj.next_question(3)
                    assess_ui_common_obj.next_question(4)
                    assess_ui_common_obj.next_question(5)
                    assess_ui_common_obj.end_test()
                    assess_ui_common_obj.end_test_confirmation()
                    self.browser.quit()

                elif current_excel_data.get('reloginRequird') == 'Yes':
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid1'))
                    assess_ui_common_obj.next_question(2)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid2'))
                    assess_ui_common_obj.next_question(3)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid3'))
                    assess_ui_common_obj.next_question(4)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid4'))
                    assess_ui_common_obj.next_question(5)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid5'))
                    self.browser.quit()
                    time.sleep(10)
                    self.browser = assess_ui_common_obj.initiate_browser(self.url, self.path)
                    login_details = assess_ui_common_obj.ui_login_to_test(current_excel_data.get('loginName'),
                                                                          (current_excel_data.get('password')))
                    if login_details == 'SUCCESS':
                        i_agreed = assess_ui_common_obj.select_i_agree()
                        if i_agreed:
                            start_test_status = assess_ui_common_obj.start_test_button_status()
                            assess_ui_common_obj.start_test()
                            answered_status_for_q1 = assess_ui_common_obj.check_answered_status(
                                current_excel_data.get('ans_qid1'))
                            print(answered_status_for_q1)
                            assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('relogin_qid1'))
                            assess_ui_common_obj.next_question(2)
                            answered_status_for_q2 = assess_ui_common_obj.check_answered_status(
                                current_excel_data.get('ans_qid1'))
                            print(answered_status_for_q2)
                            assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('relogin_qid2'))
                            assess_ui_common_obj.next_question(3)
                            answered_status_for_q3 = assess_ui_common_obj.check_answered_status(
                                current_excel_data.get('ans_qid1'))
                            print(answered_status_for_q3)
                            assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('relogin_qid3'))
                            assess_ui_common_obj.next_question(4)
                            answered_status_for_q4 = assess_ui_common_obj.check_answered_status(
                                current_excel_data.get('ans_qid1'))
                            print(answered_status_for_q4)
                            assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('relogin_qid4'))
                            assess_ui_common_obj.next_question(5)
                            answered_status_for_q5 = assess_ui_common_obj.check_answered_status(
                                current_excel_data.get('ans_qid1'))
                            print(answered_status_for_q5)
                            assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('relogin_qid5'))
                            # assess_ui_common_obj.end_test()
                            # assess_ui_common_obj.end_test_confirmation()

                else:
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid1'))
                    assess_ui_common_obj.next_question(2)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid2'))
                    assess_ui_common_obj.next_question(3)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid3'))
                    assess_ui_common_obj.next_question(4)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid4'))
                    assess_ui_common_obj.next_question(5)
                    assess_ui_common_obj.select_answer_for_the_question(current_excel_data.get('ans_qid5'))
                    # assess_ui_common_obj.end_test()
                    # assess_ui_common_obj.end_test_confirmation()
        else:
            print("login failed due to below reason")
            print(login_details)


assessment_obj = OnlineAssessment()
input_file_path = "C:/Users/User/Desktop/Automation/PythonWorkingScripts_InputData/UI/Assessment/ui_relogin.xls"
excel_read_obj.excel_read(input_file_path, 0)
excel_data = excel_read_obj.details
for current_excel_row in excel_data:
    assessment_obj.mcq_assessment(current_excel_row)
