from SCRIPTS.UI_SCRIPTS.assessment_ui_common_v2 import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.write_excel_new import *
from SCRIPTS.COMMON.io_path import *
from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.CRPO_COMMON.credentials import *


class QPVerification:

    def __init__(self):
        self.row = 1
        write_excel_object.save_result(output_path_ui_mcq_client_test_random)
        header = ['QP_Verification']
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        header = ["Test Cases", "Status", "Test Id", "Candidate Id", "Testuser ID", "User Name", "Password",
                  "mismatched questions(in QP) - question out of the QP", "Mismatched S1G1 if any",
                  "Mismatched S2G1 if any", "Mismatched S1G2 if any",
                  "Mismatched Group1 if any",
                  "Mismatched Group2 if any", "Expected Overall Randomization", "Actual Overall Randomization",
                  "Expected S1G1 Randomization", "Actual S1G1 Randomization",
                  "Expected S2G1 Randomization", "Actual S2G1 Randomization",
                  "Expected S1G2 Randomization", "Actual S1G2 Randomization", "Expected Group1 Randomization",
                  "Actual Group1 Randomization", "Expected Group2 Randomization",
                  "Actual Group2 Randomization", "Expected Group Randomization ( section position swap )",
                  "Actual Group Randomization ( section position swap )",
                  "Expected Test level Randomization ( group position swap)",
                  "Actual test level Randomization ( group position swap )",
                  "1st login Q1", "2nd login Q1",
                  "1st login Q2",
                  "2nd login Q2", "1st login Q3", "2nd login Q3", "1st login Q4", "2nd login Q4", "1st login Q5",
                  "2nd login Q5", "1st login Q6", "2nd login Q6", "1st login Q7", "2nd login Q7", "1st login Q8",
                  "2nd login Q8", "1st login Q9", "2nd login Q9", "1st login Q10", "2nd login Q10", "1st login Q11",
                  "2nd login Q11", "1st login Q12", "2nd login Q12", "1st login Q13", "2nd login Q13", "1st login Q14",
                  "2nd login Q14", "1st login Q15", "2nd login Q15", "1st login Q16", "2nd login Q16", "1st login Q17",
                  "2nd login Q17", "1st login Q18", "2nd login Q18", "1st login Q19", "2nd login Q19", "1st login Q20",
                  "2nd login Q20", "1st login Q21", "2nd login Q21"]
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

        self.qp_qn_index = [{"qn": "MS Question Randomization Low 1", "index": 1},
                            {"qn": "MS Question Randomization Low 2", "index": 2},
                            {"qn": "MS Question Randomization Low 3", "index": 3},
                            {"qn": "MS Question Randomization medium11", "index": 4},
                            {"qn": "MS Question Randomization medium21", "index": 5},
                            {"qn": "MS Question Randomization medium31", "index": 6},
                            {"qn": "MS Question Randomization high1", "index": 7},
                            {"qn": "MS Question Randomization high2", "index": 8},
                            {"qn": "MS Question Randomization high3", "index": 9},
                            {"qn": "MS Question Randomization Low 4", "index": 10},
                            {"qn": "MS Question Randomization Low 5", "index": 11},
                            {"qn": "MS Question Randomization medium4", "index": 12},
                            {"qn": "MS Question Randomization medium5", "index": 13},
                            {"qn": "MS Question Randomization high4", "index": 14},
                            {"qn": "MS Question Randomization high5", "index": 15},
                            {"qn": "MS Question Randomization Low 6", "index": 16},
                            {"qn": "MS Question Randomization Low 7", "index": 17},
                            {"qn": "MS Question Randomization high6", "index": 18},
                            {"qn": "MS Question Randohighmization 7", "index": 19},
                            {"qn": "MS Question Randomization medium6", "index": 20},
                            {"qn": "MS Question Randomization medium7", "index": 21}]

        self.section1_group1_questions = ["MS Question Randomization Low 1", "MS Question Randomization Low 2",
                                          "MS Question Randomization Low 3", "MS Question Randomization medium11",
                                          "MS Question Randomization medium21", "MS Question Randomization medium31",
                                          "MS Question Randomization high1", "MS Question Randomization high2",
                                          "MS Question Randomization high3"]

        self.section2_group1_questions = ["MS Question Randomization Low 4", "MS Question Randomization Low 5",
                                          "MS Question Randomization medium4", "MS Question Randomization medium5",
                                          "MS Question Randomization high4", "MS Question Randomization high5"]

        self.section1_group2_questions = ["MS Question Randomization Low 6", "MS Question Randomization Low 7",
                                          "MS Question Randomization high6", "MS Question Randomization high7",
                                          "MS Question Randomization medium6", "MS Question Randomization medium7"]

        self.group1_questions = self.section1_group1_questions + self.section2_group1_questions
        self.group2_questions = self.section1_group2_questions

        self.expected_questions = self.section1_group1_questions \
                                  + self.section2_group1_questions + self.section1_group2_questions

    def is_randomized(self, qn_details):
        for all_qns_with_index in self.qp_qn_index:
            if qn_details.get('question') == all_qns_with_index.get("qn"):
                if qn_details.get('index') != all_qns_with_index.get("index"):
                    self.overall_randomization = "Yes"
                    if 1 <= qn_details.get('index') <= 15:
                        if qn_details.get('question') in self.group1_questions:
                            self.g1_randomization = "Yes"
                        if 1 <= qn_details.get('index') <= 9:
                            if qn_details.get('question') in self.section1_group1_questions:
                                self.s1g1_randomization = "Yes"
                        elif 10 <= qn_details.get('index') <= 15:
                            if qn_details.get('question') in self.section2_group1_questions:
                                self.s2g1_randomization = "Yes"
                    elif 16 <= qn_details.get('index') <= 21:
                        if qn_details.get('question') in self.group2_questions:
                            self.g2_randomization = "Yes"
                        if qn_details.get('question') in self.section1_group2_questions:
                            self.s1g2_randomization = "Yes"
                break

    def verify_questions(self, tu_details, login_user, login_pass, candidate_id, tu_id):
        self.overall_randomization = "No"
        self.test_level_randomization = "No"
        self.group_randomization = "No"
        self.s1g1_randomization = "No"
        self.s2g1_randomization = "No"
        self.s1g2_randomization = "No"
        self.g1_randomization = "No"
        self.g2_randomization = "No"
        self.color = write_excel_object.green_color
        self.status = 'pass'
        self.row = self.row + 1
        self.delivered_questions = []
        self.relogin_questions = []
        self.actual_g1_questions = []
        self.actual_g2_questions = []
        self.actual_s1g1_questions = []
        self.actual_s2g1_questions = []
        self.actual_s1g2_questions = []
        self.browser = assess_ui_common_obj.initiate_browser(amsin_at_assessment_url, chrome_driver_path)
        login_details = assess_ui_common_obj.ui_login_to_test(login_user, login_pass)
        if login_details == 'SUCCESS':
            i_agreed = assess_ui_common_obj.select_i_agree()
            if i_agreed:
                start_test_status = assess_ui_common_obj.start_test_button_status()
                assess_ui_common_obj.start_test()
                for question_index in range(1, int(tu_details.get('expectedTotalQuestionsCount') + 1)):
                    assess_ui_common_obj.next_question(question_index)
                    qn_string = assess_ui_common_obj.find_question_string1()
                    print(qn_string)
                    self.delivered_questions.append(qn_string[0])
                    self.qn_details = {'question': qn_string[0], 'group': qn_string[1], 'section': qn_string[2],
                                       'index': question_index}
                    client_side_randomization.is_randomized(self.qn_details)
                    if self.qn_details.get('group') == 'Group1':
                        self.actual_g1_questions.append(self.qn_details.get('question'))
                        if self.qn_details.get('section') == "Group1Section 1":
                            self.actual_s1g1_questions.append(self.qn_details.get('question'))
                        elif self.qn_details.get('section') == "Group1Section2":
                            self.actual_s2g1_questions.append(self.qn_details.get('question'))
                        if self.qn_details.get('section') == 'Group1Section 1':
                            if 10 <= self.qn_details.get('index') <= 15:
                                self.group_randomization = "Yes"
                        elif self.qn_details.get('section') == 'Group1Section2':
                            if 1 <= self.qn_details.get('index') <= 9:
                                self.group_randomization = "Yes"
                        if 16 <= self.qn_details.get('index') <= 21:
                            self.test_level_randomization = "Yes"

                    elif self.qn_details.get('group') == 'Group2':
                        self.actual_g2_questions.append(self.qn_details.get('question'))
                        if 1 <= self.qn_details.get('index') <= 15:
                            self.test_level_randomization = "Yes"
                        if 16 <= self.qn_details.get('index') <= 21:
                            self.group_randomization = "Yes"
                        if self.qn_details.get('section') == "Group2Section 1":
                            self.actual_s1g2_questions.append(self.qn_details.get('question'))

                self.overall_randomization_status = client_side_randomization.is_randomized(self.qn_details)
                print("1st Login")
                self.browser.close()
                self.browser.switch_to.window(self.browser.window_handles[0])
                time.sleep(1)
                self.browser.execute_script(
                    '''window.open("https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0=", "_blank");''')
                time.sleep(2)
                self.browser.switch_to.window(self.browser.window_handles[1])
                login_details = assess_ui_common_obj.ui_login_to_test(login_user, login_pass)
                if login_details == 'SUCCESS':
                    i_agreed = assess_ui_common_obj.select_i_agree()
                    if i_agreed:
                        start_test_status = assess_ui_common_obj.start_test_button_status()
                        assess_ui_common_obj.start_test()
                        for question_index in range(1, int(tu_details.get('expectedTotalQuestionsCount') + 1)):
                            assess_ui_common_obj.next_question(question_index)
                            qn_string = assess_ui_common_obj.find_question_string()
                            self.relogin_questions.append(qn_string[0])
                self.browser.quit()
                write_excel_object.compare_results_and_write_vertically(tu_details.get('testCases'), None, self.row, 0)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('testID'), None, self.row, 2)
                write_excel_object.compare_results_and_write_vertically(candidate_id, None, self.row, 3)
                write_excel_object.compare_results_and_write_vertically(tu_id, None, self.row, 4)
                write_excel_object.compare_results_and_write_vertically(login_user, None, self.row, 5)
                write_excel_object.compare_results_and_write_vertically(login_pass, None, self.row, 6)

                mismatched_questions = set(self.delivered_questions) - set(self.expected_questions)
                if len(mismatched_questions) >= 1:
                    write_excel_object.current_status = 'Fail'
                    write_excel_object.overall_status = 'Fail'
                    write_excel_object.current_status_color = write_excel_object.red_color
                    write_excel_object.overall_status_color = write_excel_object.red_color
                else:
                    mismatched_questions = 'No Mismatch - Expected delivery'
                write_excel_object.ws.write(self.row, 7, str(mismatched_questions),
                                            write_excel_object.current_status_color)

                print(self.actual_s1g1_questions)
                s1g1_mismatched = set(self.section1_group1_questions) - set(self.actual_s1g1_questions)
                if len(s1g1_mismatched) >= 1:
                    write_excel_object.current_status = 'Fail'
                    write_excel_object.overall_status = 'Fail'
                    write_excel_object.current_status_color = write_excel_object.red_color
                    write_excel_object.overall_status_color = write_excel_object.red_color
                else:
                    s1g1_mismatched = 'No Mismatch - Expected delivery'
                write_excel_object.ws.write(self.row, 8, str(s1g1_mismatched),
                                            write_excel_object.current_status_color)

                s2g1_mismatched = set(self.section2_group1_questions) - set(self.actual_s2g1_questions)
                print(self.actual_s2g1_questions)
                if len(s2g1_mismatched) >= 1:
                    write_excel_object.current_status = 'Fail'
                    write_excel_object.overall_status = 'Fail'
                    write_excel_object.current_status_color = write_excel_object.red_color
                    write_excel_object.overall_status_color = write_excel_object.red_color
                else:
                    s2g1_mismatched = 'No Mismatch - Expected delivery'
                write_excel_object.ws.write(self.row, 9, str(s2g1_mismatched),
                                            write_excel_object.current_status_color)

                s1g2_mismatched = set(self.section1_group2_questions) - set(self.actual_s1g2_questions)
                print(self.actual_s1g2_questions)
                if len(s1g2_mismatched) >= 1:
                    write_excel_object.current_status = 'Fail'
                    write_excel_object.overall_status = 'Fail'
                    write_excel_object.current_status_color = write_excel_object.red_color
                    write_excel_object.overall_status_color = write_excel_object.red_color
                else:
                    s1g2_mismatched = 'No Mismatch - Expected delivery'
                write_excel_object.ws.write(self.row, 10, str(s1g2_mismatched), write_excel_object.current_status_color)

                group1_mismatched = set(self.group1_questions) - set(self.actual_g1_questions)
                if len(group1_mismatched) >= 1:
                    write_excel_object.current_status = 'Fail'
                    write_excel_object.overall_status = 'Fail'
                    write_excel_object.current_status_color = write_excel_object.red_color
                    write_excel_object.overall_status_color = write_excel_object.red_color
                else:
                    group1_mismatched = 'No Mismatch - Expected delivery'
                write_excel_object.ws.write(self.row, 11, str(group1_mismatched),
                                            write_excel_object.current_status_color)

                group2_mismatched = set(self.group2_questions) - set(self.actual_g2_questions)
                if len(group2_mismatched) >= 1:
                    write_excel_object.current_status = 'Fail'
                    write_excel_object.overall_status = 'Fail'
                    write_excel_object.current_status_color = write_excel_object.red_color
                    write_excel_object.overall_status_color = write_excel_object.red_color
                else:
                    group2_mismatched = 'No Mismatch - Expected delivery'
                write_excel_object.ws.write(self.row, 12, str(group2_mismatched),
                                            write_excel_object.current_status_color)

                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedOverallRandomization'),
                                                                        self.overall_randomization, self.row, 13)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedS1G1Randomization'),
                                                                        self.s1g1_randomization, self.row, 15)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedS2G1Randomization'),
                                                                        self.s2g1_randomization, self.row, 17)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedS1G2Randomization'),
                                                                        self.s1g2_randomization, self.row, 19)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedGroup1Randomization'),
                                                                        self.g1_randomization, self.row, 21)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedGroup2Randomization'),
                                                                        self.g2_randomization, self.row, 23)
                write_excel_object.compare_results_and_write_vertically(tu_details.get('expectedGroupRandomization'),
                                                                        self.group_randomization, self.row, 25)
                write_excel_object.compare_results_and_write_vertically(
                    tu_details.get('expectedTestLevelRandomization'),
                    self.test_level_randomization, self.row, 27)

                col = 29
                for index in range(0, int(tu_details.get('expectedTotalQuestionsCount'))):
                    print(index)
                    write_excel_object.compare_results_and_write_vertically(self.delivered_questions[index],
                                                                            self.relogin_questions[index], self.row,
                                                                            col)
                    col = col + 2
                write_excel_object.compare_results_and_write_vertically(write_excel_object.current_status, None,
                                                                        self.row, 1)


client_side_randomization = QPVerification()
excel_read_obj.excel_read(input_path_ui_mcq_client_section_random, 4)
candidate_details = excel_read_obj.details
token = crpo_common_obj.login_to_crpo(cred_crpo_admin_at.get('user'), cred_crpo_admin_at.get('password'),
                                      cred_crpo_admin_at.get('tenant'))
test_id = 15240
event_id = 11614
jobrole_id = 30434
sprint_id = input('Enter Sprint ID ')
next_cand = 2000
for current_excel_row in candidate_details:
    next_cand = next_cand + 1
    sprint_id = sprint_id + str(next_cand)
    print(sprint_id)
    candidate_id = crpo_common_obj.create_candidate(token, sprint_id)
    tag_candidate = crpo_common_obj.tag_candidate_to_test(token, candidate_id, test_id, event_id, jobrole_id)
    time.sleep(5)
    test_userid = crpo_common_obj.get_all_test_user(token, candidate_id)
    print(test_userid)
    tu_cred = crpo_common_obj.test_user_credentials(token, test_userid)
    login_id = tu_cred['data']['testUserCredential']['loginId']
    password = tu_cred['data']['testUserCredential']['password']
    client_side_randomization.verify_questions(current_excel_row, login_id, password, candidate_id, test_userid)
write_excel_object.write_overall_status(1)
