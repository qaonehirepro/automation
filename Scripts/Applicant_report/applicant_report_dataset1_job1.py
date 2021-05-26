import ast
import subprocess
from Config_dataset1_job1 import *
#----------------------------------------------------------------------------------------------------------------------#
# "ast" package is used to convert Dictionary to Json (which is used in "downloadReport" method)
# subprocess is used download the file from the webserver (which is used in "downloadReport" method)
# Config is our file which is having our configurations
#----------------------------------------------------------------------------------------------------------------------#

class AssessmentInterviewReport():
    def __init__(self):
        pass
        # print config_obj.getall_applicant_api
    # -----------------------------------------------------------------------------------------------------------------#
    #Assessment Interview Report GetallApplicant Method, API and Request data's are in Config File
    # -----------------------------------------------------------------------------------------------------------------#

    def getallApplicant(self, jobrole_id):
        config_obj.getall_applicant_data['jsonFilter']['jobIds'] = [jobrole_id]
        # config_obj.getall_applicant_data['jobRoleId'] = jobrole_id
        getall_applicant_request = requests.post(config_obj.getall_applicant_api,
                                                 headers=config_obj.headers,
                                                 data=json.dumps(config_obj.getall_applicant_data,
                                                                 default=str), verify=False)
        resp_dict = json.loads(getall_applicant_request.content)
        self.getall_applicant_status = resp_dict['status']
        self.context_guid = resp_dict['data']['ContextId']
        print "Context GUID is:- %s" % self.context_guid
        # print self.getall_applicant_status

    # -----------------------------------------------------------------------------------------------------------------#
    # Assessment Interview Report API apiStatus, This method is used to find the current Job status
    # Taken Context guid from Getall Export property response
    # -----------------------------------------------------------------------------------------------------------------#

    def apiStatus(self,contextguid):
        self.get_api_status_data = {"ContextGUID": contextguid}
        get_api_status_request = requests.post(config_obj.get_api_status,
                                               headers=config_obj.headers,
                                               data=json.dumps(self.get_api_status_data, default=str),
                                               verify=False)
        self.resp_dict = json.loads(get_api_status_request.content)
        self.get_api_status = self.resp_dict['status']
        self.get_api_job_status = self.resp_dict['data']['JobState']

    # -----------------------------------------------------------------------------------------------------------------#
    # This method downloads the Excel sheet and keeps in the user specified path
    # subprocess is a builtin method is used to download the file from the website
    # ast.literal_eval is converts dictionary type value into json type
    # -----------------------------------------------------------------------------------------------------------------#
    def downloadReport(self,jobrole_id1):
        try:
            jobrole_id = jobrole_id1
            self.getallApplicant(jobrole_id)
            if self.getall_applicant_status=='OK':
                self.get_api_job_status = 'PROGRESS'
                while self.get_api_job_status == 'PROGRESS' or self.get_api_job_status =='PENDING':
                    self.apiStatus(self.context_guid)
                    time.sleep(5)
                    print "--------------------------------"
                    print 'job Status is in %s' % self.get_api_job_status
                    print time.time()
                    print "--------------------------------"
                else:
                    print "Job status changed to Success"
                    # print self.resp_dict['data']
                    # print self.resp_dict['data']['Result']
                    downloadurl = ast.literal_eval(self.resp_dict['data']['Result'])
                    print config_obj.download_path
                    subprocess.check_output(['wget', '-O', config_obj.download_path, downloadurl['downloadLink']])
                    # print config_obj.download_path
            else:
                print "Getall Applicant Api Failed"
        except Exception as e:
            print e
    # -----------------------------------------------------------------------------------------------------------------#
    # This method reads User expected excel sheet data
    # -----------------------------------------------------------------------------------------------------------------#

    def excelReadExpectedSheet(self):
        self.expected_excel = xlrd.open_workbook(config_obj.expected_excel_sheet_path)
        self.expected_excel_sheet1 = self.expected_excel.sheet_by_index(0)

    # -----------------------------------------------------------------------------------------------------------------#
    # This method reads Downloaded excel sheet(Actual excel Sheet) data
    # -----------------------------------------------------------------------------------------------------------------#

    def excelReadActualSheet(self):
        self.actual_excel = xlrd.open_workbook(config_obj.download_path)
        self.actual_excel_sheet1 = self.actual_excel.sheet_by_index(0)

    # -----------------------------------------------------------------------------------------------------------------#
    # This method reads the data from excel sheet and writes only the headers
    # 1st for loop is used for writing first two headers (hierarchy)
    # second for loop is used for writing headers after matching
    # -----------------------------------------------------------------------------------------------------------------#

    def excelWriteHeaders(self):
        config_obj.ws.write(0, 0, "Hierarchy Header 1",config_obj.green_color)
        config_obj.ws.write(1, 0, "Hierarchy Header 2",config_obj.green_color)
        config_obj.ws.write(2, 0, "Expected header",config_obj.green_color)
        config_obj.ws.write(3, 0, "Actual header",config_obj.green_color)
        config_obj.ws.write(0, 1, "Status", config_obj.green_color)
        for i in range(0,2):
            expected_sheet_rows = air.expected_excel_sheet1.row_values(i)
            for j in range(0,air.expected_excel_sheet1.ncols):
                config_obj.ws.write(i, j+2, expected_sheet_rows[j],config_obj.green_color)

        actual_sheet_rows = air.actual_excel_sheet1.row_values(2)
        expected_sheet_rows = air.expected_excel_sheet1.row_values(2)
        for k in range(0, air.actual_excel_sheet1.ncols):
            self.status = 'Pass'
            self.color = config_obj.black_color_bold
            if expected_sheet_rows[k]==actual_sheet_rows[k]:
                config_obj.ws.write(2, k + 2, expected_sheet_rows[k],config_obj.black_color_bold)
                config_obj.ws.write(3, k + 2, actual_sheet_rows[k],config_obj.black_color_bold)
            else:
                config_obj.ws.write(2, k + 2, expected_sheet_rows[k],config_obj.red_color)
                config_obj.ws.write(3, k + 2, actual_sheet_rows[k],config_obj.red_color)
                self.status = 'Fail'
                self.color = config_obj.red_color

            config_obj.ws.write(3, 1, self.status, self.color)
    # -----------------------------------------------------------------------------------------------------------------#
    # This method matches the data between two excels and writes in the new workbook
    # workbook creation and cell formatting are there in config file
    # workbook will be created after calling config_obj.write_excel.close() method
    # -----------------------------------------------------------------------------------------------------------------#

    def excelMatchValues(self):
        self.write_position = 2

        for row_indx in range(3, air.expected_excel_sheet1.nrows):

            expected_sheet_rows = air.expected_excel_sheet1.row_values(row_indx)
            actual_sheet_rows = air.actual_excel_sheet1.row_values(row_indx)
            self.write_position += 3
            config_obj.ws.write(self.write_position, 0, "Expected Data", config_obj.green_color)
            config_obj.ws.write(self.write_position+1, 0, "Actual Data", config_obj.green_color)
            self.status = 'Pass'
            self.color = config_obj.black_color_bold
            for col_indx in range(0, air.expected_excel_sheet1.ncols):

                if expected_sheet_rows[col_indx] == actual_sheet_rows[col_indx]:
                    config_obj.ws.write(self.write_position, col_indx+2,
                                        expected_sheet_rows[col_indx], config_obj.black_color)
                    config_obj.ws.write(self.write_position+1, col_indx+2,
                                        actual_sheet_rows[col_indx], config_obj.black_color)
                else:
                    config_obj.ws.write(self.write_position, col_indx+2,
                                        expected_sheet_rows[col_indx], config_obj.red_color)
                    config_obj.ws.write(self.write_position + 1, col_indx+2,
                                        actual_sheet_rows[col_indx], config_obj.red_color)
                    self.status = 'Fail'
                    self.color = config_obj.red_color

                config_obj.ws.write(self.write_position, 1, self.status, self.color)

        # print config_obj.now
        config_obj.write_excel.close()

air = AssessmentInterviewReport()
total_jobs_count =config_obj.total_jobs.values()

for jobrole_name, jobrole_id in config_obj.total_jobs.iteritems():
    try:
        print "------------------------------------------------------------------------"
        print "Job Name :- %s     and      Job Id:- %s" %(jobrole_name,jobrole_id)
        print "------------------------------------------------------------------------"
        config_obj.filePath(jobrole_name, jobrole_id)
        config_obj.writeExcelConfigurations()
        air.downloadReport(jobrole_id)
        air.excelReadExpectedSheet()
        air.excelReadActualSheet()
        air.excelWriteHeaders()
        air.excelMatchValues()
    except Exception as e:
        print "------------------------------------------------------------------------"
        print "Job Name :- %s and Job Id:- %s is not ran Successfully, " \
              "Please verify it manually" % (jobrole_name, jobrole_id)
        print "------------------------------------------------------------------------"
        print e

