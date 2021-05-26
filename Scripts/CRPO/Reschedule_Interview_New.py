import time
import json
import requests
import xlwt
import xlrd
import datetime
import exceptions
import mysql.connector

class RescheduleInterview:

    def __init__(self):
        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------
        # Candidate_ids = '1262663,1262660,1262659,1262658,1262657,1262656'
        try:
            self.header = {"content-type": "application/json"}
            # self.TenantAlias = raw_input('TenantAlias:: ')
            # self.LoginName = raw_input('LoginName:: ')
            # self.Password = raw_input('Password:: ')

            self.login_request = {"LoginName": 'admin',
                                  "Password": '4LWS-067',
                                  "TenantAlias": "Automation",
                                  "UserName": 'admin'}
            # self.server = raw_input('Server:: ')

            login_api = requests.post("https://amsin.hirepro.in/py/common/user/login_user/",
                                      headers=self.header,
                                      data=json.dumps(self.login_request),
                                      verify=False)
            self.response = login_api.json()
            self.get_token = {"content-type": "application/json",
                              "X-AUTH-TOKEN": self.response.get("Token")}
            self.var = None
            time.sleep(1)
            resp_dict = json.loads(login_api.content)
            self.status = resp_dict['status']
            if self.status == 'OK':
                self.login = 'OK'
                print "Login successfully"
                print "Status is", self.status
                time.sleep(1)
            else:
                self.login = 'KO'
                print "Failed to login"
                print "Status is", self.status
        except exceptions.ValueError as login_error:
            print(login_error)

        # -------------------------------------------------------
        # Styles for Excel sheet Row, Column, Text - color, Font
        # -------------------------------------------------------
        self.__style0 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                    'font: name Arial, color black, bold on;')
        self.__style1 = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;'
                                    'font: name Arial, color black, bold off;')
        self.__style2 = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                                    'font: name Arial, color yellow, bold on;')
        self.__style3 = xlwt.easyxf('font: name Arial, color red, bold on')
        self.__style4 = xlwt.easyxf('pattern: pattern solid, fore_colour indigo;'
                                    'font: name Arial, color gold, bold on;')
        self.__style5 = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;'
                                    'font: name Arial, color brown, bold on;')
        self.__style6 = xlwt.easyxf('font: name Arial, color light_orange, bold on')
        self.__style7 = xlwt.easyxf('font: name Arial, color orange, bold on')
        self.__style8 = xlwt.easyxf('font: name Arial, color green, bold on')
        self.__style9 = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;'
                                    'font: name Arial, color brown, bold on;')
        self.__style10 = xlwt.easyxf('pattern: pattern solid, fore_colour brown;'
                                     'font: name Arial, color yellow, bold on;')

        # -------------------------------------
        # Excel sheet write for Output results
        # -------------------------------------
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Reschedule')
        self.rowsize = 1
        self.size = self.rowsize
        self.col = 0

        index = 0
        excelheaders = ['Comparision', 'Status', 'Message', 'Interview_Request', 'Interviewers',
                        'interview_type', 'ScheduleOn', 'interviewStatus', 'Reschedule_ir_id',
                        'ReSchedule_Interviewers', 'ReSchedule_type', 'ReScheduleOn', 'Reschedule_Status',
                        'ApplicantId']
        for headers in excelheaders:
            if headers in ['Comparision', 'Message', 'Status']:
                self.ws.write(0, index, headers, self.__style10)
            elif headers in ['Interview_Request', 'Interviewers', 'interview_type', 'ScheduleOn', 'interviewStatus']:
                self.ws.write(0, index, headers, self.__style2)
            elif headers in ['Reschedule_ir_id', 'add_interviewers', 'Removed_interviewers', 'ReSchedule_type',
                             'Reschedule_Status', 'ReScheduleOn', 'ReSchedule_Interviewers']:
                self.ws.write(0, index, headers, self.__style9)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1
        print ('Excel Headers are printed successfully')

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_Event_id = []  # [] Initialising data from excel sheet to the variables
        self.xl_Applicant_id = []
        self.xl_Job_id = []
        self.xl_type = []
        self.xl_Schedule_Datetime = []
        self.xl_stage_id = []
        self.xl_interviewers_id = []
        self.xl_Schedule_Comment = []
        self.xl_location = []

        # --------------------------
        # Reschedule and cancel data
        # --------------------------
        self.xl_Reschedule_DateTime = []
        self.xl_Reschedule_type = []
        self.xl_Reschedule_add_interviewers = []
        self.xl_Reschedule_remove_interviewers = []
        self.xl_Reschedule_comment = []

        # -----------------------------------------------------------------------------------
        # Dictionaries for Interview_schedule, interview_feedback, interview_feedback_details
        # -----------------------------------------------------------------------------------
        self.ir = {}
        self.i_r = self.ir
        self.reschedule_ir = {}
        self.r_ir = self.reschedule_ir

        self.is_success = {}
        self.i_s = self.is_success
        self.is_reschedule_success = {}
        self.i_s = self.is_reschedule_success

        self.message = {}
        self.msg = self.message
        self.reschedule_message = {}
        self.msg1 = self.reschedule_message

        self.candidate_details_dict = {}
        self.cdd = self.candidate_details_dict
        self.applicant_details_dict = {}
        self.add = self.applicant_details_dict
        self.data = {}
        self.d = self.data
        self.interviewers = {}
        self.int = self.interviewers

        self.updated_candidate_details_dict = {}
        self.u_cdd = self.updated_candidate_details_dict
        self.updated_applicant_details_dict = {}
        self.u_add = self.updated_applicant_details_dict
        self.updated_data = {}
        self.ud = self.updated_data
        self.updated_interviewers = {}
        self.u_i = self.updated_interviewers

        self.final_status = {}
        self.f_s = self.final_status

    def ams_connection(self):
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()
        self.ids = self.xl_Applicant_id
        self.ids = tuple(self.ids)
        self.ir_query = "select ir.id from candidates c " \
                        "inner join interview_candidates ic on c.id=ic.candidate_id " \
                        "inner join interview_requests ir on ic.interviewrequest_id=ir.id " \
                        "where ic.applicant_status_id in{} " \
                        "and c.tenant_id=1787 and ir.is_deleted=0;".format(self.ids)
        # print self.ir_query
        # query = self.ir_query
        time.sleep(2)
        self.cursor.execute(self.ir_query)
        interview_request_ids = self.cursor.fetchall()
        self.cancellation_ids = []
        for ids in interview_request_ids:
            self.cancellation_ids.append(ids[0])

    def cancel_interview(self):
        cancel_request = {"interviewRequestIds": self.cancellation_ids, "interviewCanceledStatusId": 167112,
                          "applicantStatusItemComment": "cancel_request"}

        ir_details_api = requests.post('https://amsin.hirepro.in/py/crpo/api/v1/interview/cancel/',
                                       headers=self.get_token,
                                       data=json.dumps(cancel_request, default=str), verify=False)
        # cancel_interview_response = json.loads(ir_details_api.content)
        # print  ir_details_api
        # data = cancel_interview_response
        # print data



    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_InputData/'
                                          'CRPO/Reschedule Interview/Reschedule.xls')
            sheet = workbook.sheet_by_index(0)
            for i in range(1, sheet.nrows):
                number = i
                rows = sheet.row_values(number)

                if rows[0] is not None and rows[0] != '':
                    self.xl_Event_id.append(int(rows[0]))
                else:
                    self.xl_Event_id.append(None)

                if rows[1] is not None and rows[1] != '':
                    self.xl_Applicant_id.append(int(rows[1]))
                else:
                    self.xl_Applicant_id.append(None)

                if rows[2] is not None and rows[2] != '':
                    self.xl_Job_id.append(int(rows[2]))
                else:
                    self.xl_Job_id.append(None)

                if rows[3] is not None and rows[3] != '':
                    self.xl_type.append(int(rows[3]))
                else:
                    self.xl_type.append(None)

                if rows[4] is not None and rows[4] != '':
                    self.xl_Schedule_Datetime.append(str(rows[4]))
                else:
                    self.xl_Schedule_Datetime.append(None)

                if rows[5] is not None and rows[5] != '':
                    self.xl_stage_id.append(int(rows[5]))
                else:
                    self.xl_stage_id.append(None)

                if rows[6] is not None and rows[6] != '':
                    int_ids = map(int, rows[6].split(',') if isinstance(rows[6], basestring) else [rows[6]])
                    self.xl_interviewers_id.append(int_ids)
                else:
                    self.xl_interviewers_id.append(None)

                if rows[7] is not None and rows[7] != '':
                    self.xl_Schedule_Comment.append(str(rows[7]))
                else:
                    self.xl_Schedule_Comment.append(None)

                if rows[8] is not None and rows[8] != '':
                    self.xl_location.append(int(rows[8]))
                else:
                    self.xl_location.append(None)

                if rows[9] is not None and rows[9] != '':
                    self.xl_Reschedule_DateTime.append(str(rows[9]))
                else:
                    self.xl_Reschedule_DateTime.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_Reschedule_type.append(int(rows[10]))
                else:
                    self.xl_Reschedule_type.append(None)

                if rows[11] is not None and rows[11] != '':
                    int_ids = map(int, rows[11].split(',') if isinstance(rows[11], basestring) else [rows[11]])
                    self.xl_Reschedule_add_interviewers.append(int_ids)
                else:
                    self.xl_Reschedule_add_interviewers.append(None)

                if rows[12] is not None and rows[12] != '':
                    int_ids = map(int, rows[12].split(',') if isinstance(rows[12], basestring) else [rows[12]])
                    self.xl_Reschedule_remove_interviewers.append(int_ids)
                else:
                    self.xl_Reschedule_remove_interviewers.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_Reschedule_comment.append(str(rows[13]))
                else:
                    self.xl_Reschedule_comment.append(None)

            print('Excel data initiated is Done')

        except IOError:
            print("File not found or path is incorrect")

    def schedule_interview(self, loop):
        try:
            schedule_request = [{
                "isConsultantRound": False,
                "interviewDate": self.xl_Schedule_Datetime[loop],
                "interviewTime": "",
                "interviewType": self.xl_type[loop],
                "interviewerIds": self.xl_interviewers_id[loop],
                "jobId": self.xl_Job_id[loop],
                "stageId": self.xl_stage_id[loop],
                "locationId": self.xl_location[loop],  # API default send bangalore location
                "secondaryInterviewerIds": [],
                "recruiterComment": self.xl_Schedule_Comment[loop],
                "recruitEventId": self.xl_Event_id[loop],
                "applicantIds": [self.xl_Applicant_id[loop]]
            }]
            scheduling_interviews = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/interview/schedule/",
                                                  headers=self.get_token,
                                                  data=json.dumps(schedule_request, default=str), verify=False)
            schedule_response = json.loads(scheduling_interviews.content)
            # print (json.dumps(schedule_response, indent=2))
            data = schedule_response['data']
            # print(json.dumps(data, indent=2))
            # print('***--------------------------------------------------------***')

            if schedule_response['status'] == 'OK':
                success = data['success']
                failure = data['failure']

                if data['success']:
                    for i in success:
                        self.ir = i['interviewRequestId']
                        print self.ir
                        print "Scheduled to interview"
                        self.message = i.get('message')
                        if self.xl_Applicant_id[loop] is not None:
                            self.is_success = True
                elif data['failure']:
                    for i in failure:
                        self.message = i.get('message')
                        print self.message
                        self.is_success = False
            else:
                print ('Error occured while scheduling')
        except exceptions.ValueError as Schedule_error:
            print(Schedule_error)

    def reschedule_interview(self, loop):
        reschedule_request = [{
            "interviewRequestId": self.ir,
            "interviewType": self.xl_Reschedule_type[loop],
            "interviewers": self.xl_Reschedule_add_interviewers[loop],
            "interviewDate": self.xl_Reschedule_DateTime[loop],
            "recruiterComment": self.xl_Reschedule_comment[loop],
            "removedInterviewers": self.xl_Reschedule_remove_interviewers[loop]
        }]

        reschedule_api = requests.post('https://amsin.hirepro.in/py/crpo/api/v1/interview/reschedule/',
                                       headers=self.get_token,
                                       data=json.dumps(reschedule_request, default=str), verify=False)
        reschedule_response = json.loads(reschedule_api.content)
        data = reschedule_response['data']

        if reschedule_response['status'] == 'OK':
            success = data['success']
            failure = data['failure']
            print failure

            if data['success']:
                for i in success:
                    self.reschedule_ir = i['interviewId']
                    print self.reschedule_ir
                    print "ReScheduled to interview"
                    if self.xl_Applicant_id[loop] is not None:
                        self.is_reschedule_success = True
            elif data['failure']:
                for i in failure:
                    self.reschedule_message = i.get('message')
                    print self.message
                    self.is_reschedule_success = False

    def interview_request_details(self, loop):
        ir_details_request = {
            "search": {
                "interviewrequests": [self.reschedule_ir]
            }}
        ir_details_api = requests.post('https://amsin.hirepro.in/py/crpo/api/v1/view/interviews',
                                       headers=self.get_token,
                                       data=json.dumps(ir_details_request, default=str), verify=False)
        ir_details_response = json.loads(ir_details_api.content)
        data = ir_details_response['data']
        for i in data:
            self.data = i
            self.candidate_details_dict = i.get('candidate')
            self.applicant_details_dict = i.get('applicant')

            interviewer_details_dict = i.get('interviewers')
            int = interviewer_details_dict.keys()
            self.interviewers = ', '.join(int)

    def updated_interview_request_details(self, loop):
        ir_details_request = {
            "search": {
                "interviewrequests": [self.ir]
            }}
        ir_details_api = requests.post('https://amsin.hirepro.in/py/crpo/api/v1/view/interviews',
                                       headers=self.get_token,
                                       data=json.dumps(ir_details_request, default=str), verify=False)
        ir_details_response = json.loads(ir_details_api.content)
        data = ir_details_response['data']
        for i in data:
            self. updated_data = i
            self.updated_candidate_details_dict = i.get('candidate')
            self.updated_applicant_details_dict = i.get('applicant')

            updated_interviewer_details_dict = i.get('interviewers')
            updated_int = updated_interviewer_details_dict.keys()
            self.updated_interviewers = ', '.join(updated_int)
            self.final_status = True

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
        self.ws.write(self.rowsize, 4, str(self.xl_interviewers_id[loop]), self.__style1)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_type[loop] == 2:
            self.ws.write(self.rowsize, 5, 'FaceToFace', self.__style1)
        elif self.xl_type[loop] == 3:
            self.ws.write(self.rowsize, 5, 'WebCam', self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 6, self.xl_Schedule_Datetime[loop], self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 9, str(self.xl_Reschedule_add_interviewers[loop]), self.__style1)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_Reschedule_type[loop] == 2:
            self.ws.write(self.rowsize, 10, 'FaceToFace', self.__style1)
        elif self.xl_Reschedule_type[loop] == 3:
            self.ws.write(self.rowsize, 10, 'WebCam', self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 11, self.xl_Reschedule_DateTime[loop], self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 13, self.xl_Applicant_id[loop], self.__style1)
        # --------------------------------------------------------------------------------------------------------------

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
        # --------------------------------------------------------------------------------------------------------------

        if self.final_status:
            self.ws.write(self.rowsize, 1, 'True', self.__style8)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.__style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.message:
            if 'Applicant has been already' in self.message:
                self.ws.write(self.rowsize, 2, self.message, self.__style3)
            else:
                self.ws.write(self.rowsize, 2, self.message, self.__style8)
        elif self.reschedule_message:
            self.ws.write(self.rowsize, 2, self.reschedule_message, self.__style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.ir:
            self.ws.write(self.rowsize, 3, self.ir, self.__style8)
        else:
            self.ws.write(self.rowsize, 3, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.interviewers:
            if self.xl_interviewers_id[loop] is None:
                self.ws.write(self.rowsize, 4, 'NoInterviewers', self.__style7)
            else:
                self.ws.write(self.rowsize, 4, self.interviewers, self.__style8)
        else:
            self.ws.write(self.rowsize, 4, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.data and self.data.get('typeOfInterview'):
            if self.xl_type[loop] == self.data.get('typeOfInterviewEnum'):
                if self.xl_type[loop] is None:
                    self.ws.write(self.rowsize, 5, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 5, self.data.get('typeOfInterview'), self.__style8)
            else:
                self.ws.write(self.rowsize, 5, None)
        else:
            self.ws.write(self.rowsize, 5, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.data and self.data.get('scheduledOn'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 6, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 6, self.data.get('scheduledOn'), self.__style8)
            else:
                self.ws.write(self.rowsize, 6, None)
        else:
            self.ws.write(self.rowsize, 6, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_details_dict and self.applicant_details_dict.get('currentStatus'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 7, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 7, self.applicant_details_dict.get('currentStatus'), self.__style8)
            else:
                self.ws.write(self.rowsize, 7, None)
        else:
            self.ws.write(self.rowsize, 7, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.reschedule_ir:
            self.ws.write(self.rowsize, 8, self.reschedule_ir, self.__style8)
        else:
            self.ws.write(self.rowsize, 8, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_interviewers:
            if self.xl_Reschedule_add_interviewers is None:
                self.ws.write(self.rowsize, 9, 'NoInterviewers', self.__style7)
            else:
                self.ws.write(self.rowsize, 9, self.updated_interviewers, self.__style8)
        else:
            self.ws.write(self.rowsize, 9, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_data and self.updated_data.get('typeOfInterview'):
            if self.xl_Reschedule_type[loop] == self.updated_data.get('typeOfInterviewEnum'):
                if self.xl_Reschedule_type[loop] is None:
                    self.ws.write(self.rowsize, 10, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 10, self.updated_data.get('typeOfInterview'), self.__style8)
            else:
                self.ws.write(self.rowsize, 10, None)
        else:
            self.ws.write(self.rowsize, 10, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_data and self.updated_data.get('scheduledOn'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 11, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 11, self.updated_data.get('scheduledOn'), self.__style8)
            else:
                self.ws.write(self.rowsize, 11, None)
        else:
            self.ws.write(self.rowsize, 11, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_applicant_details_dict and self.updated_applicant_details_dict.get('currentStatus'):
            if self.xl_Applicant_id[loop] == self.updated_applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 12, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 12, self.updated_applicant_details_dict.get('currentStatus'),
                                  self.__style8)
            else:
                self.ws.write(self.rowsize, 12, None)
        else:
            self.ws.write(self.rowsize, 12, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_details_dict and self.applicant_details_dict.get('applicantId'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id is None:
                    self.ws.write(self.rowsize, 13, 'Empty_value', self.__style8)
                else:
                    self.ws.write(self.rowsize, 13, self.applicant_details_dict.get('applicantId'))
            else:
                self.ws.write(self.rowsize, 13, None)
        else:
            self.ws.write(self.rowsize, 13, None)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Object.wb_Result.save('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/CRPO/'
                              'Reschedule Interview/API_Reschedule.xls')



Object = RescheduleInterview()
Object.excel_data()
Object.ams_connection()
Object.cancel_interview()
Total_count = len(Object.xl_Event_id)
print "Number of Rows ::", Total_count

try:
    if Object.login == 'OK':
        for looping in range(0, Total_count):
            print "Iteration Count is ::", looping
            Object.schedule_interview(looping)
            Object.reschedule_interview(looping)
            if Object.is_reschedule_success:
                Object.interview_request_details(looping)
                Object.updated_interview_request_details(looping)
                Object.interview_request_details(looping)
            Object.output_excel(looping)
            print('Excel data is ready')

            # -------------------------------------
            # Making all dict empty for every loop
            # -------------------------------------
            Object.is_success = {}
            Object.is_reschedule_success = {}
            Object.ir = {}
            Object.reschedule_ir = {}
            Object.message = {}
            Object.reschedule_message = {}

            Object.candidate_details_dict = {}
            Object.applicant_details_dict = {}
            Object.interviewers = {}

            Object.updated_candidate_details_dict = {}
            Object.updated_applicant_details_dict = {}
            Object.updated_interviewers = {}

            Object.data = {}
            Object.updated_data = {}

            Object.final_status = {}

except exceptions.AttributeError as Object_error:
    print(Object_error)
