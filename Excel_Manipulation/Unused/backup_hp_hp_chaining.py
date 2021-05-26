import xlsxwriter
import datetime
from COMMON.read_excel import *
from Unused.assessment_common import *


class ChainingOfTwoTests:

    def __init__(self):
        self.started = datetime.datetime.now()
        self.started = self.started.strftime("%Y-%m-%d-%H-%M-%S")
        # self.row_size = 2
        self.write_excel = xlsxwriter.Workbook(
            'C:\\Users\\User\Desktop\\Automation\\PythonWorkingScripts_Output'
            '\\Assessment\\HP_HP_Chaining_Automation_Tenant - ' + self.started + '.xls')

        self.ws = self.write_excel.add_worksheet()
        self.black_color = self.write_excel.add_format({'font_color': 'black', 'font_size': 9})
        self.red_color = self.write_excel.add_format({'font_color': 'red', 'font_size': 9})
        self.green_color = self.write_excel.add_format({'font_color': 'green', 'font_size': 9})
        self.orange_color = self.write_excel.add_format({'font_color': 'orange', 'font_size': 9})
        self.black_color_bold = self.write_excel.add_format({'font_color': 'black', 'bold': True, 'font_size': 9})
        self.over_all_status_pass = self.write_excel.add_format({'font_color': 'green', 'bold': True, 'font_size': 9})
        self.over_all_status_failed = self.write_excel.add_format({'font_color': 'red', 'bold': True, 'font_size': 9})
        self.over_all_status_color = self.over_all_status_pass
        self.over_all_status = 'OVERALL STATUS :- Pass'
        self.ws.write(0, 0, "HP HP Chaining", self.black_color_bold)
        self.ws.write(1, 0, "Testcases", self.black_color_bold)
        self.ws.write(1, 1, "Testcase Status", self.black_color_bold)
        self.ws.write(1, 2, "Test ID", self.black_color_bold)
        self.ws.write(1, 3, "Candidate ID", self.black_color_bold)
        self.ws.write(1, 4, "Login Name", self.black_color_bold)
        self.ws.write(1, 5, "Password", self.black_color_bold)
        self.ws.write(1, 6, "First test is type of ", self.black_color_bold)
        self.ws.write(1, 7, "Is auto test", self.black_color_bold)
        self.ws.write(1, 8, "Is SLC Enabled", self.black_color_bold)
        self.ws.write(1, 9, "T1 Expected Status", self.black_color_bold)
        self.ws.write(1, 10, "T1 Actual Status", self.black_color_bold)

        self.ws.write(1, 11, "Expected - are you expecting next test info?", self.black_color_bold)
        self.ws.write(1, 12, "Actual - Is next test info available in response?", self.black_color_bold)
        self.ws.write(1, 13, "if next test is available, Are you able to login to next test?", self.black_color_bold)
        self.ws.write(1, 14, "Expected domain URL for next test", self.black_color_bold)
        self.ws.write(1, 15, "Actual domain URL for next test", self.black_color_bold)
        self.ws.write(1, 16, "Expected candidate id in second test ", self.black_color_bold)
        self.ws.write(1, 17, "Actual candidate id in second test ", self.black_color_bold)
        self.ws.write(1, 18, "Expected second test id", self.black_color_bold)
        self.ws.write(1, 19, "Actual second test id", self.black_color_bold)
        self.ws.write(1, 20, "Expected second test type", self.black_color_bold)
        self.ws.write(1, 21, "Actual second test type", self.black_color_bold)

    def compare_data(self, expected_data, actual_data):
        if expected_data == actual_data:
            self.color = self.green_color

        else:
            self.color = self.red_color
            self.actual_status = 'Fail'
            self.over_all_status_color = self.red_color
            self.actual_status_color = self.red_color
            self.over_all_status = 'OVERALL STATUS :- Fail'

    def final_report(self, row_value, next_test_info, excel_data, next_tu_infos):
        self.actual_status = 'Pass'
        self.actual_status_color = self.green_color

        self.compare_data(next_test_info.get('actualStatus'), excel_data.get('expectedStatus'))
        self.ws.write(row_value, 9, excel_data.get('expectedStatus'), self.color)
        self.ws.write(row_value, 10, next_test_info.get('actualStatus'), self.color)

        self.compare_data(excel_data.get('expectingNextTest'), next_test_info.get('nextTestAvailability'))
        self.ws.write(row_value, 11, excel_data.get('expectingNextTest'), self.color)
        self.ws.write(row_value, 12, next_test_info.get('nextTestAvailability'), self.color)

        self.compare_data(excel_data.get('expectedDomainUrlForSecondTest'), next_test_info.get('nextTestDomainHost'))
        self.ws.write(row_value, 14, excel_data.get('expectedDomainUrlForSecondTest'), self.color)
        self.ws.write(row_value, 15, next_test_info.get('nextTestDomainHost'), self.color)

        if excel_data.get('expectedCandidateId') != 'EMPTY':
            expected_candidate_id_in_next_test = int(excel_data.get('expectedCandidateId'))
        else:
            expected_candidate_id_in_next_test = excel_data.get('expectedCandidateId')

        self.compare_data(expected_candidate_id_in_next_test, next_tu_infos.get('nextTestCandidateId'))
        self.ws.write(row_value, 16, expected_candidate_id_in_next_test, self.color)
        self.ws.write(row_value, 17, next_tu_infos.get('nextTestCandidateId'), self.color)

        if excel_data.get('expectedSecondTestId') != 'EMPTY':
            expected_test_id_in_next_test = int(excel_data.get('expectedSecondTestId'))
        else:
            expected_test_id_in_next_test = excel_data.get('expectedSecondTestId')

        self.compare_data(expected_test_id_in_next_test, next_tu_infos.get('nextTestID'))
        self.ws.write(row_value, 18, expected_test_id_in_next_test, self.color)
        self.ws.write(row_value, 19, next_tu_infos.get('nextTestID'), self.color)

        self.compare_data(excel_data.get('expectedSecondTestType'), next_tu_infos.get('next_test_type'))
        self.ws.write(row_value, 20, excel_data.get('expectedSecondTestType'), self.color)
        self.ws.write(row_value, 21, next_tu_infos.get('next_test_type'), self.color)

        self.ws.write(row_value, 0, excel_data.get('testCaseInfo'), self.black_color)
        self.ws.write(row_value, 1, self.actual_status, self.actual_status_color)
        self.ws.write(row_value, 2, excel_data.get('primaryTestId'), self.black_color)
        self.ws.write(row_value, 3, excel_data.get('candidateId'), self.black_color)
        self.ws.write(row_value, 4, excel_data.get('loginName'), self.black_color)
        self.ws.write(row_value, 5, excel_data.get('password'), self.black_color)
        self.ws.write(row_value, 6, excel_data.get('firstTestIsTypeOf'), self.black_color)
        self.ws.write(row_value, 7, excel_data.get('isAutoTest'), self.black_color)
        self.ws.write(row_value, 8, excel_data.get('IsSLCEnabled'), self.black_color)
        self.ws.write(row_value, 13, next_tu_infos.get('login_staus'), self.black_color)


chaining_obj = ChainingOfTwoTests()
input_file_path = 'C:\\Users\\User\\Desktop\\Automation\\PythonWorkingScripts_InputData\\Assessment\\chaining\\chaining_of_2_tests_automation_tenant.xls'

excel_read_obj.excel_read(input_file_path, 0)
excel_data = excel_read_obj.details
crpo_headers = crpo_common_obj.login_to_crpo(login_name='rpm', password='Rpmuthu@123', tenant='automation')
test_id = 10190
candidate_ids = [1301162, 1301164, 1301163, 1301165]
crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)

test_id = 10192
candidate_ids = [1301162, 1301164, 1301163, 1301165]
crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)

test_id = 10199
candidate_ids = [1301162]
crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)

test_id = 10201
candidate_ids = [1301166]
crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)

test_id = 10209
candidate_ids = [1301162]
crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)

test_id = 10206
candidate_ids = [1301162]
crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)
# test_id = 10047
# candidate_ids = [1299331, 1299329, 1299328, 1299327]
# crpo_common_obj.untag_candidate(crpo_headers, test_id, candidate_ids)
rowsize = 1
for data in excel_data:
    domain = data.get('firstTestDomain')
    rowsize = rowsize + 1
    login_details = assessment_common_obj.login_to_test(login_name=data.get('loginName'),
                                                        password=data.get('password'),
                                                        tenant='automation', type_of_test=data.get('firstTestIsTypeOf'))
    submit_token = assessment_common_obj.submit_test_result(assessment_token=login_details.get('login_token'),
                                                            domain=login_details.get('domain'),
                                                            submit_test_request=data.get('submitTestVariableName'))
    if submit_token:
        if data.get('firstTestIsTypeOf') == 'VET' and data.get('IsCallbackRequired') == 'Yes':
            print( data.get('scoreCallBack'))
            print( type( data.get('scoreCallBack')))
            assessment_common_obj.pearson_call_backs(int(data.get('testUserId')), data.get('scoreCallBack'))

        next_test_details = {}
        next_tu_infos1 = {}
        actual_status = "None"
        initiate_automation_resp = assessment_common_obj.initiate_automation(submit_token, data.get('candidateId'),
                                                                             int(data.get('primaryTestId')),
                                                                             domain=login_details.get('domain'))
        initiate_automation_data = initiate_automation_resp.get('data')
        if initiate_automation_data.get('contextId'):
            print("this is chaining test")
            polling_api_response = assessment_common_obj.get_job_status(submit_token,
                                                                        initiate_automation_data.get('contextId'))
            if polling_api_response['data']['JobState'] == "SUCCESS":
                # next_test_details = assessment_common_obj.process_next_test_link_for_slc_test(status[0])
                next_test_details = assessment_common_obj.process_next_test_links_for_chaining(polling_api_response,
                                                                                               domain)
                print(next_test_details)

                if next_test_details.get('actualStatus') == 'Shortlisted' or next_test_details.get(
                        'actualStatus') == 'EMPTY':
                    next_tu_infos1 = assessment_common_obj.login_to_test(
                        login_name=next_test_details.get('next_test_login_id'),
                        password=next_test_details.get('next_test_pwd'),
                        tenant='automation',
                        type_of_test=data.get('firstTestIsTypeOf'))
                    next_tu_infos1 = next_tu_infos1.get('test_user_infos')
                else:
                    next_tu_infos1 = {'nextTestID': 'EMPTY',
                                      'nextTestName': 'EMPTY',
                                      'nextTestCandidateId': 'EMPTY', 'login_staus': 'EMPTY', 'next_test_type': 'EMPTY'}
        else:
            print("this is not a chaining test")
            next_tu_infos1 = {'nextTestID': 'EMPTY', 'nextTestName': 'EMPTY', 'nextTestCandidateId': 'EMPTY',
                              'login_staus': 'EMPTY', 'next_test_type': 'EMPTY'}
        chaining_obj.final_report(rowsize, next_test_details, data, next_tu_infos1)
    else:
        print("Submit token is not available")

ended = datetime.datetime.now()
ended = "Ended:- %s" % ended.strftime("%Y-%m-%d-%H-%M-%S")
chaining_obj.ws.write(0, 1, chaining_obj.over_all_status, chaining_obj.over_all_status_color)
chaining_obj.ws.write(0, 2, 'Started:- ' + chaining_obj.started, chaining_obj.black_color_bold)
chaining_obj.ws.write(0, 3, ended, chaining_obj.black_color_bold)
chaining_obj.ws.write(0, 4, "Total_Testcase_Count:- %s" % len(excel_data), chaining_obj.black_color_bold)
chaining_obj.write_excel.close()
