import xlsxwriter
from Excel_Manipulation.COMMON.read_excel import *
# return requests.get(url).json()
import datetime
import time
from  Excel_Manipulation.CRPO.credentials import *
from  Excel_Manipulation.CRPO.crpo_common import *
from  Excel_Manipulation.ASSESSMENT.assessment_common import *


class ReInitiateAutomation:
    def __init__(self):
        # requests.packages.urllib3.disable_warnings()
        self.started = datetime.datetime.now()
        self.started = self.started.strftime("%Y-%M-%d-%H-%M-%S")
        self.row_size = 2
        self.write_excel = xlsxwriter.Workbook(
            'D:\\automation\\PythonWorkingScripts_Output'
            '\\Assessment\\reinitiate\\reinitiate - ' + self.started + '.xls')
        self.final_status = ''
        self.ws = self.write_excel.add_worksheet()
        self.black_color = self.write_excel.add_format({'font_color': 'black', 'font_size': 9})
        self.red_color = self.write_excel.add_format({'font_color': 'red', 'font_size': 9})
        self.green_color = self.write_excel.add_format({'font_color': 'green', 'font_size': 9})
        self.orange_color = self.write_excel.add_format({'font_color': 'orange', 'font_size': 9})
        self.black_color_bold = self.write_excel.add_format({'font_color': 'black', 'bold': True, 'font_size': 9})
        self.over_all_status_pass = self.write_excel.add_format({'font_color': 'green', 'bold': True, 'font_size': 9})
        self.over_all_status_failed = self.write_excel.add_format({'font_color': 'red', 'bold': True, 'font_size': 9})
        self.over_all_status_color = self.over_all_status_pass
        self.over_all_status = 'Pass'
        self.ws.write(0, 0, "Reinitiate relogin automation", self.black_color_bold)
        self.ws.write(1, 0, "Testcases", self.black_color_bold)
        self.ws.write(1, 1, "Status", self.black_color_bold)
        self.ws.write(1, 2, "Test ID", self.black_color_bold)
        self.ws.write(1, 3, "Candidate ID", self.black_color_bold)
        self.ws.write(1, 4, "Login Name", self.black_color_bold)
        self.ws.write(1, 5, "Password", self.black_color_bold)
        self.ws.write(1, 6, "Is Vendor Test", self.black_color_bold)
        self.ws.write(1, 7, "Is SLC Enabled", self.black_color_bold)
        self.ws.write(1, 8, "Expected Status", self.black_color_bold)
        self.ws.write(1, 9, "Actual Status", self.black_color_bold)

    def test_user_next_test_status(self, login_response):
        print(login_response)
        if login_response.get('status') == 'KO':
            next_test_flags = login_response.get('error').get('nextTestFlags')
            if next_test_flags:
                if next_test_flags.get('isScoreFetched'):
                    if next_test_flags.get('isShortlisted'):
                        if next_test_flags.get('isRegistered'):
                            print("Candidate is registered in the vendor test")
                            self.final_status = "Shortlisted"
                        else:
                            print("Candidate is not registered in the vendor test")
                            self.final_status = "Rejected"
                    else:
                        print("Score is available but not shortlisted for the next test")
                        self.final_status = "Rejected"
                else:
                    if next_test_flags.get('isHproTest') == True:
                        print("Chaining test without SLC")
                        if next_test_flags.get('isRegistered') == False:
                            self.final_status = "Rejected"
                        else:
                            self.final_status = "Shortlisted"
                    else:
                        print("Score is not available")
                        self.final_status = "Rejected"
            else:
                print("Test is already submitted")
                self.final_status = "Completed"
        else:
            print("First test is not completed")
            self.final_status = "Pending"

    def excel_write(self, data):

        if self.final_status == data.get('expectedStatus'):
            status = "Pass"
            final_status_color = self.green_color
        else:
            status = "Fail"
            self.over_all_status = "Fail"
            self.over_all_status_color = self.red_color
            final_status_color = self.red_color
        self.ws.write(self.row_size, 0, data.get('testCaseInfo'), self.black_color)
        self.ws.write(self.row_size, 1, status, final_status_color)
        self.ws.write(self.row_size, 2, data.get('primaryTestId'), self.black_color)
        self.ws.write(self.row_size, 3, data.get('candidateId'), self.black_color)
        self.ws.write(self.row_size, 4, data.get('loginName'), self.black_color)
        self.ws.write(self.row_size, 5, data.get('password'), self.black_color)
        self.ws.write(self.row_size, 6, data.get('isVendorTest'), self.black_color)
        self.ws.write(self.row_size, 7, data.get('IsSLCEnabled'), self.black_color)
        self.ws.write(self.row_size, 8, data.get('expectedStatus'), final_status_color)
        self.ws.write(self.row_size, 9, self.final_status, final_status_color)
        self.row_size = self.row_size + 1


re_initiate_obj = ReInitiateAutomation()
input_file_path = 'D:\\automation\\PythonWorkingScripts_InputData\\Assessment' \
                  '\\reinitiateautomation1.xls'
excel_read_obj.excel_read(input_file_path, 0)
excel_data = excel_read_obj.details
crpo_headers = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                             cred_crpo_admin.get('tenant'))

untag_candidates_details = [{"testUserIds": [893441, 893442, 893443]},
                            {"testUserIds": [893444]},
                            {"testUserIds": [893445]},
                            {"testUserIds": [893446, 893447, 893448, 893449]}]
crpo_common_obj.untag_candidate(crpo_headers, untag_candidates_details)

for data in excel_data:
    print(data.get('loginName'), data.get('password'))
    test_id = int(data.get('primaryTestId'))
    candiate_id = int(data.get('candidateId'))
    crpo_common_obj.re_initiate_automation(crpo_headers, test_id, candiate_id)
    time.sleep(5)
    assessment_headers = AssessmentCommon.login_to_test(login_name=data.get('loginName'), password=data.get('password'),
                                                        tenant='automation')

    re_initiate_obj.test_user_next_test_status(assessment_headers)
    re_initiate_obj.excel_write(data)

ended = datetime.datetime.now()
ended = "Ended:- %s" % ended.strftime("%Y-%M-%d-%H-%M-%S")
re_initiate_obj.ws.write(0, 1, re_initiate_obj.over_all_status, re_initiate_obj.over_all_status_color)
re_initiate_obj.ws.write(0, 2, 'Started:- ' + re_initiate_obj.started, re_initiate_obj.black_color_bold)
re_initiate_obj.ws.write(0, 3, ended, re_initiate_obj.black_color_bold)
re_initiate_obj.ws.write(0, 4, "Total_Testcase_Count:- 29", re_initiate_obj.black_color_bold)
re_initiate_obj.write_excel.close()
