from SCRIPTS.UI_COMMON.assessment_ui_common_v2 import *
from SCRIPTS.UI_SCRIPTS.assessment_data_verification import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.writeExcel import *
from SCRIPTS.COMMON.io_path import *


class TestSecurity:

    def __init__(self):
        self.row = 1
        write_excel_object.save_result(output_path_ui_test_security)
        self.url = "https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF1dG9tYXRpb24ifQ=="
        # self.path = r"F:\qa_automation\chromedriver.exe"
        header = ['QP_Verification']
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        self.overall_test_case_status = 'Pass'
        self.overall_test_case_status_color = write_excel_object.green_color
        header = ['Test Cases', 'Status', 'Test Id', 'Candidate Id', 'Testuser ID', 'User Name', 'Password',
                  'Expected Status', 'Actual Status']
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

    def test_security(self, candidate_details, test_user_info):
        self.row = self.row + 1
        delivered_questions = []
        self.browser = assess_ui_common_obj.initiate_browser(amsin_automation_assessment_url, chrome_driver_path)
        login_details = assess_ui_common_obj.ui_login_to_test(candidate_details.get('userName'),
                                                              candidate_details.get('password'))
        if login_details == 'SUCCESS':
            i_agreed = assess_ui_common_obj.select_i_agree()
            if i_agreed:
                start_test_status = assess_ui_common_obj.start_test_button_status()
                assess_ui_common_obj.start_test()
                time.sleep(2)
                secure_password_model_window = assess_ui_common_obj.check_security_key_model_window_availability()
                if secure_password_model_window == 'Success':
                    assess_ui_common_obj.validate_security_key(candidate_details.get('securityKey'))
                secure_password_model_window1 = assess_ui_common_obj.check_security_key_model_window_availability()
                if secure_password_model_window1 == 'Success':
                    print("Invalid Password")
                    actual_status = 'Invalid Password'
                else:
                    print("Valid Psssword")
                    actual_status = 'Valid Password'

                write_excel_object.ws.write(self.row, 0, candidate_details.get('testCases'),
                                            write_excel_object.green_color)
                write_excel_object.ws.write(self.row, 2, candidate_details.get('testID'),
                                            write_excel_object.green_color)
                write_excel_object.ws.write(self.row, 3, candidate_details.get('candidateID'),
                                            write_excel_object.green_color)
                write_excel_object.ws.write(self.row, 4, candidate_details.get('testUserId'),
                                            write_excel_object.green_color)
                write_excel_object.ws.write(self.row, 5, candidate_details.get('userName'),
                                            write_excel_object.green_color)
                write_excel_object.ws.write(self.row, 6, candidate_details.get('password'),
                                            write_excel_object.green_color)
                write_excel_object.ws.write(self.row, 7, candidate_details.get('expectedStatus'),
                                            write_excel_object.green_color)
                if candidate_details.get('expectedStatus') == actual_status:
                    write_excel_object.ws.write(self.row, 8, actual_status, write_excel_object.green_color)
                    testcase_status = 'Pass'
                    testcase_status_color = write_excel_object.green_color
                else:
                    write_excel_object.ws.write(self.row, 8, actual_status, write_excel_object.red_color)
                    testcase_status = 'Fail'
                    self.overall_test_case_status = 'Fail'
                    testcase_status_color = write_excel_object.red_color
                    self.overall_test_case_status_color = write_excel_object.red_color
                write_excel_object.ws.write(self.row, 1, testcase_status, testcase_status_color)

        self.browser.close()


print(datetime.datetime.now())
test_security_obj = TestSecurity()
excel_read_obj.excel_read(input_path_ui_test_security, 0)
excel_data = excel_read_obj.details
# print(excel_data)
# assessment_obj.verify_questions(candidate_details, questions)
for current_excel_row in excel_data:
    print(current_excel_row)
    test_security_obj.test_security(current_excel_row, excel_data)

write_excel_object.ws.write(0, 1, test_security_obj.overall_test_case_status,
                            test_security_obj.overall_test_case_status_color)
write_excel_object.write_excel.close()
# crpo_token = crpo_common_obj.login_to_crpo('admin', 'Email@admin', 'AT')
# time.sleep(10)
# obj_assessment_data_verification.assessment_data_report(crpo_token, excel_data)
# obj_assessment_data_verification.write_excel.close()
# print(datetime.datetime.now())
