from automation.Excel_Manipulation.uiautomation.assessment_ui_common_v2 import *
from automation.Excel_Manipulation.uiautomation.assessment_data_verification import *
from automation.Excel_Manipulation.COMMON.read_excel import *
from automation.Excel_Manipulation.COMMON.writeExcel import *


class QPVerification:

    def __init__(self):
        self.row = 1

        save_path = r"F:\qa_automation\automation\PythonWorkingScripts_Output\UI\QP_"
        write_excel_object.save_result(save_path)
        self.url = "https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0="
        self.path = r"F:\qa_automation\automation\chromedriver.exe"
        header = ['QP_Verification']
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        header = ['Test Cases', 'Status', 'Test Id', 'Candidate Id', 'Testuser ID', 'User Name', 'Password',
                  'Expected qns', 'Actual qns ', 'mismatched questions - which is not part of expected if any']
        self.expected_questions = ['MS Question Randomization Low 1', 'MS Question Randomization Low 2',
                                   'MS Question Randomization Low 3', 'MS Question Randomization Low 4',
                                   'MS Question Randomization Low 5',
                                   'MS Question Randomization medium1', 'MS Question Randomization medium2',
                                   'MS Question Randomization medium3', 'MS Question Randomization medium4',
                                   'MS Question Randomization medium5', 'MS Question Randomization high1',
                                   'MS Question Randomization high2', 'MS Question Randomization high3',
                                   'MS Question Randomization high4', 'MS Question Randomization high5']
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

    def verify_questions(self, candidate_details, qn_infos):
        color = write_excel_object.green_color
        status = 'pass'
        self.row = self.row + 1
        delivered_questions = []
        self.browser = assess_ui_common_obj.initiate_browser(self.url, self.path)
        login_details = assess_ui_common_obj.ui_login_to_test(candidate_details.get('userName'),
                                                              candidate_details.get('password'))
        if login_details == 'SUCCESS':
            i_agreed = assess_ui_common_obj.select_i_agree()
            if i_agreed:
                start_test_status = assess_ui_common_obj.start_test_button_status()
                assess_ui_common_obj.start_test()
                qn_string = assess_ui_common_obj.find_question_string()
                delivered_questions.append(qn_string)
                assess_ui_common_obj.next_question(2)
                qn_string = assess_ui_common_obj.find_question_string()
                delivered_questions.append(qn_string)
                assess_ui_common_obj.next_question(3)
                qn_string = assess_ui_common_obj.find_question_string()
                delivered_questions.append(qn_string)
                assess_ui_common_obj.next_question(4)
                qn_string = assess_ui_common_obj.find_question_string()
                delivered_questions.append(qn_string)
                assess_ui_common_obj.next_question(5)
                qn_string = assess_ui_common_obj.find_question_string()
                delivered_questions.append(qn_string)
                assess_ui_common_obj.next_question(6)
                qn_string = assess_ui_common_obj.find_question_string()
                delivered_questions.append(qn_string)
                print(delivered_questions)
                self.col = 7
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
                write_excel_object.ws.write(self.row, 7, str(self.expected_questions), color)
                write_excel_object.ws.write(self.row, 7, str(delivered_questions), color)

                mismatched_questions = set(delivered_questions) - set(self.expected_questions)
                print(mismatched_questions)
                if len(mismatched_questions) >= 1:
                    color = write_excel_object.red_color
                    status = 'fail'
                    overall_status = 'Fail'
                    overall_color = write_excel_object.red_color
                write_excel_object.ws.write(self.row, 9, str(mismatched_questions), color)
                write_excel_object.ws.write(self.row, 1, status, color)
                # self.col = self.col + 2

            # for excel in qn_infos:
            #     for delivered_questions in delivered_questions:
            #         if excel['questions'] in actual['questions']:
            #             write_excel_object.ws.write(self.row, self.col, excel['questions'],
            #                                         write_excel_object.green_color)
            #             write_excel_object.ws.write(self.row, self.col + 1, actual['questions'],
            #                                         write_excel_object.green_color)
            #             write_excel_object.ws.write(self.row, 1, 'Pass',
            #                                         write_excel_object.green_color)
            #             print(excel['questions'])
            #             print(actual['questions'])
            #             print('Matched')
            #             self.col = self.col + 2
            #             break
            #     else:
            #         write_excel_object.ws.write(self.row, self.col, excel['questions'],
            #                                     write_excel_object.green_color)
            #         write_excel_object.ws.write(self.row, self.col + 1, "Question Not Available",
            #                                     write_excel_object.red_color)
            #         write_excel_object.ws.write(self.row, 1, 'Fail',
            #                                     write_excel_object.green_color)
            #         self.col = self.col + 2

        self.browser.close()


print(datetime.datetime.now())
assessment_obj = QPVerification()
input_file_path = r"F:\qa_automation\automation\PythonWorkingScripts_InputData\UI\Assessment\qprandomization.xls"
excel_read_obj.excel_read(input_file_path, 0)
questions = excel_read_obj.details
# print(questions)
excel_read_obj.details = []
excel_read_obj.excel_read(input_file_path, 1)
candidate_details = excel_read_obj.details
# print(candidate_details)
# assessment_obj.verify_questions(candidate_details, questions)
for current_excel_row in candidate_details:
    # print(current_excel_row)
    assessment_obj.verify_questions(current_excel_row, questions)

write_excel_object.write_excel.close()
# crpo_token = crpo_common_obj.login_to_crpo('admin', 'Email@admin', 'AT')
# time.sleep(10)
# obj_assessment_data_verification.assessment_data_report(crpo_token, excel_data)
# obj_assessment_data_verification.write_excel.close()
# print(datetime.datetime.now())
