import xlsxwriter
import datetime
from COMMON.read_excel import *
from CRPO.crpo_common import *
from CRPO.credentials import *
from COMMON.writeExcel import *


class ProctorEvaluation:
    def __init__(self):
        save_path = "C:\\Users\\User\Desktop\\Automation\\PythonWorkingScripts_Output\\Assessment\\proctoring_evaluation"
        excel_object.save_result(save_path)
        # 0th Row Header
        header = ['Proctoring Evaluation automation']
        #1 Row Header
        excel_object.write_headers_for_scripts(0, 0, header, excel_object.black_color_bold)
        header = ['Testcases', 'Status', 'Test ID', 'Candidate ID', 'Testuser ID', 'Expected img proctoring status',
                  'Actual Img proctoring status', 'Expected Video proctoring status', 'Actual Video proctoring status',
                  'Expected Audio proctoring status', 'Actual Audio proctoring status', 'Expected overall proctoring status',
                  'Actual overall proctoring status', 'Expected overall rating', 'Actual overall rating']
        excel_object.write_headers_for_scripts(1, 0, header, excel_object.black_color_bold)

    def suspicious_or_not_supicious(self, data, overall_proctoring_status_value):
        if data is True:
            if overall_proctoring_status_value is False:
                self.status = 'Suspicious'
            else:
                if overall_proctoring_status_value >= 0:
                    self.status = 'Not Suspicious'
                elif 0 < overall_proctoring_status_value <= 0.33:
                    self.status = 'Low'
                elif 0.33 < overall_proctoring_status_value < 0.66:
                    self.status = 'Medium'
                else:
                    self.status = 'Highly Suspicious'
        else:
            self.status = 'Not Suspicious'

    def compare_results(self, excel_data, actual_data):
        if excel_data == actual_data:
            self.current_status = 'Pass'
            self.current_status_color = excel_object.green_color
        else:
            self.current_status = 'Fail'
            self.over_all_status = 'Fail'
            self.current_status_color = excel_object.red_color
            self.overalll_status_color = excel_object.red_color

    def proctor_detail(self, row_count, current_excel_data, token):
        self.over_all_status = 'Pass'
        self.overalll_status_color = excel_object.green_color
        tu_id = int(current_excel_data.get('testUserId'))
        tu_proctor_details = crpo_common_obj.proctor_evaluation_detail(token, tu_id)
        proctorDetail = tu_proctor_details['data']['proctorDetail']
        image_proctoring_status = proctorDetail.get('imgSuspicious')

        self.suspicious_or_not_supicious(image_proctoring_status, False)
        self.compare_results(current_excel_data.get('imageProctoringStatus'), self.status)
        excel_object.ws.write(row_count, 5, current_excel_data.get('imageProctoringStatus'), self.current_status_color)
        excel_object.ws.write(row_count, 6, self.status, self.current_status_color)

        video_proctoring_status = proctorDetail.get('faceSuspicious')
        self.suspicious_or_not_supicious(video_proctoring_status, False)
        self.compare_results(current_excel_data.get('videoProctoringStatus'), self.status)
        excel_object.ws.write(row_count, 7, current_excel_data.get('videoProctoringStatus'), self.current_status_color)
        excel_object.ws.write(row_count, 8, self.status, self.current_status_color)

        audio_proctoring_status = proctorDetail.get('audioSuspicious')
        self.suspicious_or_not_supicious(audio_proctoring_status, False)
        self.compare_results(current_excel_data.get('audioProctoringStatus'), self.status)
        excel_object.ws.write(row_count, 9, current_excel_data.get('audioProctoringStatus'), self.current_status_color)
        excel_object.ws.write(row_count, 10, self.status, self.current_status_color)

        overall_proctoring_status = proctorDetail.get('finalDecision')
        overall_suspicious_value = proctorDetail.get('systemOverallDecision')
        self.suspicious_or_not_supicious(overall_proctoring_status, overall_suspicious_value)
        self.compare_results(current_excel_data.get('overallProctoringStatus'), self.status)
        excel_object.ws.write(row_count, 11, current_excel_data.get('overallProctoringStatus'), self.current_status_color)
        excel_object.ws.write(row_count, 12, self.status, self.current_status_color)

        self.compare_results(current_excel_data.get('overallSuspiciousValue'), overall_suspicious_value)
        excel_object.ws.write(row_count, 13, current_excel_data.get('overallSuspiciousValue'), self.current_status_color)
        excel_object.ws.write(row_count, 14, overall_suspicious_value, self.current_status_color)

        excel_object.ws.write(row_count, 0, current_excel_data.get('testCase'), excel_object.black_color_bold)
        excel_object.ws.write(row_count, 1, self.current_status, self.current_status_color)
        excel_object.ws.write(row_count, 2, current_excel_data.get('testId'), excel_object.black_color_bold)
        excel_object.ws.write(row_count, 3, current_excel_data.get('candidateId'), excel_object.black_color_bold)
        excel_object.ws.write(row_count, 4, current_excel_data.get('testUserId'), excel_object.black_color_bold)
        self.row_count = row_count + 1


login_token = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                            cred_crpo_admin.get('tenant'))
excel_read_obj.excel_read(
    'C:\\Users\\User\\Desktop\\Automation\\PythonWorkingScripts_InputData\\Assessment\\proc_eval\\proc_eval.xls', 0)
excel_data = excel_read_obj.details
proctor_obj = ProctorEvaluation()
tuids = []
for fetch_tuids in excel_data:
    tuids.append(int(fetch_tuids.get('testUserId')))
context_id = CrpoCommon.force_evaluate_proctoring(login_token, tuids)
context_id = context_id['data']['ContextId']
print(context_id)
current_job_status = 'Pending'
while current_job_status == 'Pending':
    current_job_status = CrpoCommon.job_status(login_token, context_id)
    current_job_status = current_job_status['data']['JobState']
    print("_________________ Proctor Evaluation is in Progress _______________________")
    print(current_job_status)
for data in excel_data:
    row_count = 2
    proctor_obj.proctor_detail(row_count, data, login_token)
    row_count = row_count + 1
ended = datetime.datetime.now()
ended = "Ended:- %s" % ended.strftime("%Y-%M-%d-%H-%M-%S")
excel_object.ws.write(0, 1, proctor_obj.over_all_status, proctor_obj.overalll_status_color)
excel_object.ws.write(0, 2, 'Started:- ' + excel_object.started, excel_object.black_color_bold)
excel_object.ws.write(0, 3, ended, excel_object.black_color_bold)
excel_object.ws.write(0, 4, "Total_Testcase_Count:- 2", excel_object.black_color_bold)
excel_object.write_excel.close()
