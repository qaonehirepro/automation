import ast
from Excel_Manipulation.applicantReportConfig import *
from Excel_Manipulation.COMMON.writeExcel import *
import urllib.request
from Excel_Manipulation.CRPO.crpo_common import *
from Excel_Manipulation.CRPO.credentials import *
import time


# ----------------------------------------------------------------------------------------------------------------------#
# "ast" package is used to convert Dictionary to Json (which is used in "downloadReport" method)
# subprocess is used download the file from the webserver (which is used in "downloadReport" method)
# Config is our file which is having our configurations
# ----------------------------------------------------------------------------------------------------------------------#

class AssessmentInterviewReport:
    # -----------------------------------------------------------------------------------------------------------------#
    # Assessment Interview Report GetallApplicant Method, API and Request data's are in Config File
    # -----------------------------------------------------------------------------------------------------------------#
    # -----------------------------------------------------------------------------------------------------------------#
    # This method downloads the Excel sheet and keeps in the user specified path
    # subprocess is a builtin method is used to download the file from the website
    # ast.literal_eval is converts dictionary type value into json type
    # -----------------------------------------------------------------------------------------------------------------#
    def downloadReport(self, token):
        resp_dict = crpo_common_obj.generate_applicant_report(token, config_obj.getall_applicant_data)
        self.getall_applicant_status = resp_dict['status']
        self.context_guid = resp_dict['data']['ContextId']
        print("Context GUID is:- %s" % self.context_guid)
        if self.getall_applicant_status == 'OK':
            self.get_api_job_status = 'PROGRESS'
            while self.get_api_job_status == 'PROGRESS' or self.get_api_job_status == 'PENDING':
                self.resp_dict = crpo_common_obj.job_status(token, self.context_guid)
                self.get_api_job_status = self.resp_dict['data']['JobState']
                time.sleep(5)
                print("--------------------------------")
                print('job Status is in %s' % self.get_api_job_status)
                print(time.time())
                print("--------------------------------")
            else:
                print("Job status changed to Success")
                downloadurl = ast.literal_eval(self.resp_dict['data']['Result'])
                print(config_obj.download_path)
                urllib.request.urlretrieve(downloadurl.get('downloadLink'), config_obj.download_path)
                # urllib.request.urlretrieve(downloadurl, config_obj.download_path)

                # subprocess.check_output(['wget', '-O', config_obj.download_path, downloadurl['downloadLink']])
                # print config_obj.download_path
        else:
            print("Getall Applicant Api Failed")


air = AssessmentInterviewReport()
total_jobs_count = config_obj.total_jobs.values()
crpo_headers = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                             cred_crpo_admin.get('tenant'))
for jobrole_name, jobrole_id in config_obj.total_jobs.items():
    print("------------------------------------------------------------------------")
    print("Job Name :- %s     and      Job Id:- %s" % (jobrole_name, jobrole_id))
    print("------------------------------------------------------------------------")
    config_obj.filePath(jobrole_name, jobrole_id)
    config_obj.writeExcelConfigurations()
    air.downloadReport(crpo_headers)
    write_excel_object.save_result(config_obj.save_path)
    write_excel_object.excelReadExpectedSheet(config_obj.expected_excel_sheet_path)
    write_excel_object.excelReadActualSheet(config_obj.download_path)
    write_excel_object.excelWriteHeaders(hierarchy_headers_count=3)
    write_excel_object.excelMatchValues(usecase_name='Applicant Report', comparision_required_from_index=3,
                                        total_testcase_count=48)
    # try:
    #     print("------------------------------------------------------------------------")
    #     print("Job Name :- %s     and      Job Id:- %s" % (jobrole_name, jobrole_id))
    #     print("------------------------------------------------------------------------")
    #     config_obj.filePath(jobrole_name, jobrole_id)
    #     config_obj.writeExcelConfigurations()
    #     # air.downloadReport(crpo_headers)
    #     write_excel_object.save_result(config_obj.save_path)
    #     write_excel_object.excelReadExpectedSheet(config_obj.expected_excel_sheet_path)
    #     write_excel_object.excelReadActualSheet(config_obj.download_path)
    #     write_excel_object.excelWriteHeaders(hierarchy_headers_count=3)
    #     write_excel_object.excelMatchValues(usecase_name='Applicant Report', comparision_required_from_index=3,
    #                                         total_testcase_count=48)
    # except Exception as e:
    #     print("------------------------------------------------------------------------")
    #     print("Job Name :- %s and Job Id:- %s is not ran Successfully, " \
    #           "Please verify it manually" % (jobrole_name, jobrole_id))
    #     print("------------------------------------------------------------------------")
    #     print(e)
