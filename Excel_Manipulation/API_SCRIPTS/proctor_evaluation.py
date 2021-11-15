from automation.Excel_Manipulation.CRPO_COMMON.crpo_common import *
from automation.Excel_Manipulation.CRPO_COMMON.credentials import *
from automation.Excel_Manipulation.COMMON.read_excel import *
from automation.Excel_Manipulation.COMMON.writeExcel import *
from automation.Excel_Manipulation.COMMON.io_path import *

class ProctorEvaluation:
    def __init__(self):
        write_excel_object.save_result(output_path_proctor_evaluation)
        # 0th Row Header
        header = ['Proctoring Evaluation automation']
        # 1 Row Header
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        header = ['Testcases', 'Status', 'Test ID', 'Candidate ID', 'Testuser ID', 'Expected img proctoring status',
                  'Actual Img proctoring status', 'Expected Video proctoring status', 'Actual Video proctoring status',
                  'Expected Audio proctoring status', 'Actual Audio proctoring status',
                  'Expected overall proctoring status',
                  'Actual overall proctoring status', 'Expected overall rating', 'Actual overall rating']
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

    def suspicious_or_not_supicious(self, data, overall_proctoring_status_value):
        if data is True:
            if overall_proctoring_status_value is False:
                self.status = 'Suspicious'
            else:
                if overall_proctoring_status_value >= 0.66:
                    self.status = 'Highly Suspicious'

                elif overall_proctoring_status_value >= 0.35:
                    self.status = 'Medium'

                elif overall_proctoring_status_value > 0:
                    self.status = 'Low'
                else:
                    self.status = 'Not Suspicious'


                # if 0 < overall_proctoring_status_value <= 0.35:
                #     self.status = 'Low'
                # elif 0.35 < overall_proctoring_status_value <= 0.66:
                #     self.status = 'Medium'
                # elif 0.66 < overall_proctoring_status_value <= 1:
                #     self.status = 'Highly Suspicious'
                # else:
                #     self.status = 'Not Suspicious'
        else:
            self.status = 'Not Suspicious'

    def compare_results(self, excel_data, actual_data):
        if excel_data == actual_data:
            self.current_status_color = write_excel_object.green_color
        else:
            self.current_status = 'Fail'
            proctor_obj.over_all_status = 'Fail'
            self.current_status_color = write_excel_object.red_color
            proctor_obj.overalll_status_color = write_excel_object.red_color
            self.test_case_status_color = write_excel_object.red_color

    def proctor_detail(self, row_count, current_excel_data, token):
        self.current_status_color = write_excel_object.green_color
        self.current_status = 'Pass'
        self.test_case_status_color = write_excel_object.green_color
        tu_id = int(current_excel_data.get('testUserId'))
        tu_proctor_details = crpo_common_obj.proctor_evaluation_detail(token, tu_id)
        proctorDetail = tu_proctor_details['data']['proctorDetail']
        image_proctoring_status = proctorDetail.get('imgSuspicious')

        self.suspicious_or_not_supicious(image_proctoring_status, False)
        self.compare_results(current_excel_data.get('imageProctoringStatus'), self.status)
        write_excel_object.ws.write(row_count, 5, current_excel_data.get('imageProctoringStatus'),
                                    self.current_status_color)
        write_excel_object.ws.write(row_count, 6, self.status, self.current_status_color)

        video_proctoring_status = proctorDetail.get('faceSuspicious')
        self.suspicious_or_not_supicious(video_proctoring_status, False)
        self.compare_results(current_excel_data.get('videoProctoringStatus'), self.status)
        write_excel_object.ws.write(row_count, 7, current_excel_data.get('videoProctoringStatus'),
                                    self.current_status_color)
        write_excel_object.ws.write(row_count, 8, self.status, self.current_status_color)

        audio_proctoring_status = proctorDetail.get('audioSuspicious')
        self.suspicious_or_not_supicious(audio_proctoring_status, False)
        self.compare_results(current_excel_data.get('audioProctoringStatus'), self.status)
        write_excel_object.ws.write(row_count, 9, current_excel_data.get('audioProctoringStatus'),
                                    self.current_status_color)
        write_excel_object.ws.write(row_count, 10, self.status, self.current_status_color)

        overall_proctoring_status = proctorDetail.get('finalDecision')
        overall_suspicious_value = proctorDetail.get('systemOverallDecision')
        self.suspicious_or_not_supicious(overall_proctoring_status, overall_suspicious_value)


        self.compare_results(current_excel_data.get('overallProctoringStatus'), self.status)
        write_excel_object.ws.write(row_count, 11, current_excel_data.get('overallProctoringStatus'),
                                    self.current_status_color)
        write_excel_object.ws.write(row_count, 12, self.status, self.current_status_color)

        excel_overall_suspicious_value = round(current_excel_data.get('overallSuspiciousValue'), 2)
        self.compare_results(excel_overall_suspicious_value, overall_suspicious_value)
        write_excel_object.ws.write(row_count, 13, excel_overall_suspicious_value,
                                    self.current_status_color)
        write_excel_object.ws.write(row_count, 14, overall_suspicious_value, self.current_status_color)

        write_excel_object.ws.write(row_count, 0, current_excel_data.get('testCase'),
                                    write_excel_object.black_color_bold)
        write_excel_object.ws.write(row_count, 1, self.current_status, self.test_case_status_color)
        write_excel_object.ws.write(row_count, 2, current_excel_data.get('testId'), write_excel_object.black_color_bold)
        write_excel_object.ws.write(row_count, 3, current_excel_data.get('candidateId'),
                                    write_excel_object.black_color_bold)
        write_excel_object.ws.write(row_count, 4, current_excel_data.get('testUserId'),
                                    write_excel_object.black_color_bold)


login_token = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                            cred_crpo_admin.get('tenant'))
excel_read_obj.excel_read(input_path_proctor_evaluation, 0)
excel_data = excel_read_obj.details
proctor_obj = ProctorEvaluation()
tuids = []
overalll_status_color = write_excel_object.green_color
over_all_status = 'Pass'

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
row_count = 2
for data in excel_data:
    proctor_obj.proctor_detail(row_count, data, login_token)
    row_count = row_count + 1
ended = datetime.datetime.now()
ended = "Ended:- %s" % ended.strftime("%Y-%M-%d-%H-%M-%S")
write_excel_object.ws.write(0, 1, proctor_obj.over_all_status, proctor_obj.overalll_status_color)
write_excel_object.ws.write(0, 2, 'Started:- ' + write_excel_object.started, write_excel_object.black_color_bold)
write_excel_object.ws.write(0, 3, ended, write_excel_object.black_color_bold)
write_excel_object.ws.write(0, 4, "Total_Testcase_Count:- 2", write_excel_object.black_color_bold)
write_excel_object.write_excel.close()
