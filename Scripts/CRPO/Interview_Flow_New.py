import time
import json
import requests
import xlwt
import xlrd
import datetime
import exceptions


class InterviewFeedback:

    def __init__(self):
        try:
            self.header = {"content-type": "application/json"}
            self.login_request = {"LoginName": "Interviewer_01",
                                  "Password": "Interviewer_01",
                                  "TenantAlias": "Automation",
                                  "UserName": "Interviewer_01"}

            login_api = requests.post("https://amsin.hirepro.in/py/common/user/login_user/",
                                      headers=self.header,
                                      data=json.dumps(self.login_request),
                                      verify=False)
            self.response = login_api.json()
            self.get_token = {"content-type": "application/json",
                              "X-AUTH-TOKEN": self.response.get("Token")}
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
        self.__style0 = xlwt.easyxf('font: name Arial, color black, bold on;')
        self.__style1 = xlwt.easyxf('font: name Arial, color black, bold off;')
        self.__style2 = xlwt.easyxf('font: name Arial, color green, bold on')
        self.__style3 = xlwt.easyxf('font: name Arial, color red, bold on')
        self.__style4 = xlwt.easyxf('font: name Arial, color gold, bold on;')
        self.__style5 = xlwt.easyxf('font: name Arial, color brown, bold on;')
        self.__style6 = xlwt.easyxf('font: name Arial, color light_orange, bold on')

        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Schedule')
        self.rowsize = 1
        self.col = 0

        index = 0
        excelheaders = ['Comparison', 'Status', 'ApplicantID', 'Schedule_Message', 'IR_id', 'Feedback_Message',
                        'Interviewer_Decision', 'Partial_Feedback', 'Partial_Feedback_message',
                        'Update_decision_message', 'Updated_decision', 'Scheduled_date', 'Interviewed_date', 'Skill_01',
                        'Score_01', 'Skill_02', 'Score_02', 'Skill_03', 'Score_03', 'Skill_04', 'Score_04', 'Duration',
                        'Skill_comment', 'OverAllComment', 'Update_Score_01', 'Update_Duration',
                        'Update_Skill_comment', 'Updated_OverAllComment']
        for headers in excelheaders:
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
        self.xl_Datetime = []
        self.xl_stage_id = []
        self.xl_interviewers_id = []
        self.xl_Schedule_Comment = []
        self.xl_location = []

        self.xl_Skill_id_01 = []
        self.xl_Skill_score_01 = []
        self.xl_Skill_id_02 = []
        self.xl_Skill_score_02 = []
        self.xl_Skill_id_03 = []
        self.xl_Skill_score_03 = []
        self.xl_Skill_id_04 = []
        self.xl_Skill_score_04 = []
        self.xl_skill_comment = []
        self.xl_decision = []
        self.xl_duration = []
        self.xl_int_datetime = []
        self.xl_Over_all_comment = []
        self.xl_partial_feedback = []
        self.xl_expected_status= []

        # -----------------------------------------
        # Update details / Partial feedback details
        # -----------------------------------------
        self.xl_updated_duration = []
        self.xl_Updated_Over_all_comment = []
        self.xl_update_Skill_comment = []
        self.xl_update_skill_score_01 = []
        self.xl_update_stage = []

        # -----------------------------------------------------------------------------------
        # Dictionaries for Interview_schedule, interview_feedback, interview_feedback_details
        # -----------------------------------------------------------------------------------
        self.ir = {}
        self.is_success = {}
        self.is_feedback = {}
        self.message = {}
        self.feedback = {}
        self.feedback_data = {}
        self.updated_feedback_data = {}
        self.updated_feedback = {}
        self.partial_data = {}
        self.feedback_message = {}
        self.applicant_details = {}

        # -------------------
        # Skill dictionaries
        # -------------------
        self.skill_dict_1 = {}
        self.skill_dict_2 = {}
        self.skill_dict_3 = {}
        self.skill_dict_4 = {}
        self.filledFeedbackDetails = {}
        self.skillAssessed_details = {}

        # ---------------------------
        # Skill updated dictionaries
        # ---------------------------
        self.updated_skill_dict_1 = {}
        self.updated_skill_dict_2 = {}
        self.updated_skill_dict_3 = {}
        self.updated_skill_dict_4 = {}
        self.updated_filledFeedbackDetails = {}
        self.updated_skillAssessed_details = {}
        self.decision_update = {}
        self.updatedecision = {}
        self.decision_error = {}
        self.decision_updated_feedback = {}

        # ----------------------------
        # Partial/updated Dictionaries
        # ----------------------------
        self.pf = {}

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/'
                                          'PythonWorkingScripts_InputData/CRPO/Give Feedback/GiveFeedback.xls')
            sheet = workbook.sheet_by_index(0)
            for i in range(1, sheet.nrows):
                number = i
                rows = sheet.row_values(number)
                if rows[0]:
                    self.xl_Event_id.append(int(rows[0]))
                else:
                    self.xl_Event_id.append(None)

                if rows[1]:
                    self.xl_Applicant_id.append(int(rows[1]))
                else:
                    self.xl_Applicant_id.append(None)

                if rows[2]:
                    self.xl_Job_id.append(int(rows[2]))
                else:
                    self.xl_Job_id.append(None)

                if rows[3]:
                    self.xl_type.append(int(rows[3]))
                else:
                    self.xl_type.append(None)

                self.xl_Datetime.append(str(rows[4]))

                if rows[5]:
                    self.xl_stage_id.append(int(rows[5]))
                else:
                    self.xl_stage_id.append(None)

                if rows[6]:
                    int_ids = map(int, rows[6].split(',') if isinstance(rows[6], basestring) else [rows[6]])
                    self.xl_interviewers_id.append(int_ids)
                else:
                    self.xl_interviewers_id.append(None)

                self.xl_Schedule_Comment.append(str(rows[7]))

                if rows[8]:
                    self.xl_location.append(int(rows[8]))
                else:
                    self.xl_location.append(None)

                if rows[9]:
                    self.xl_Skill_id_01.append(int(rows[9]))
                else:
                    self.xl_Skill_id_01.append(None)

                if rows[10]:
                    self.xl_Skill_score_01.append(int(rows[10]))
                else:
                    self.xl_Skill_score_01.append(None)

                if rows[11]:
                    self.xl_Skill_id_02.append(int(rows[11]))
                else:
                    self.xl_Skill_id_02.append(None)

                if rows[12]:
                    self.xl_Skill_score_02.append(int(rows[12]))
                else:
                    self.xl_Skill_score_02.append(None)

                if rows[13]:
                    self.xl_Skill_id_03.append(int(rows[13]))
                else:
                    self.xl_Skill_id_03.append(None)

                if rows[14]:
                    self.xl_Skill_score_03.append(int(rows[14]))
                else:
                    self.xl_Skill_score_03.append(None)

                if rows[15]:
                    self.xl_Skill_id_04.append(int(rows[15]))
                else:
                    self.xl_Skill_id_04.append(None)

                if rows[16]:
                    self.xl_Skill_score_04.append(int(rows[16]))
                else:
                    self.xl_Skill_score_04.append(None)

                self.xl_skill_comment.append(str(rows[17]))

                if rows[18]:
                    self.xl_decision.append(int(rows[18]))
                else:
                    self.xl_decision.append(None)

                if rows[19]:
                    self.xl_duration.append(int(rows[19]))
                else:
                    self.xl_duration.append(None)

                self.xl_int_datetime.append(str(rows[20]))

                self.xl_Over_all_comment.append(str(rows[21]))

                if rows[22]:
                    self.xl_partial_feedback.append(int(rows[22]))
                else:
                    self.xl_partial_feedback.append(None)

                if rows[23]:
                    self.xl_updated_duration.append(int(rows[23]))
                else:
                    self.xl_updated_duration.append(None)

                self.xl_Updated_Over_all_comment.append(str(rows[24]))

                self.xl_update_Skill_comment.append(str(rows[25]))

                if rows[26]:
                    self.xl_update_skill_score_01.append(int(rows[26]))
                else:
                    self.xl_update_skill_score_01.append(None)

                if rows[27]:
                    self.xl_update_stage.append(int(rows[27]))
                else:
                    self.xl_update_stage.append(None)

                self.xl_expected_status.append(str(rows[30]))

            print('Excel data initiated is Done')

        except IOError:
            print("File not found or path is incorrect")

    def schedule_interview(self, loop):
        try:
            schedule_request = [{
                "isConsultantRound": False,
                "interviewDate": self.xl_Datetime[loop],
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

    def provide_feedback(self, loop):
        if self.xl_int_datetime[loop]:
            if self.xl_partial_feedback[loop] == 1:
                self.pf = True
            else:
                self.pf = False
            try:
                feedback_request = {
                    "interviewRequestId": self.ir,
                    "interviewerFeedback": [{
                        "partial_feedback": self.pf,
                        "skillsAssessed": [{
                            "skillId": self.xl_Skill_id_01[loop],
                            "skillScore": self.xl_Skill_score_01[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }, {
                            "skillId": self.xl_Skill_id_02[loop],
                            "skillScore": self.xl_Skill_score_02[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }, {
                            "skillId": self.xl_Skill_id_03[loop],
                            "skillScore": self.xl_Skill_score_03[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }, {
                            "skillId": self.xl_Skill_id_04[loop],
                            "skillScore": self.xl_Skill_score_04[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }],
                        "interviwerIds": self.xl_interviewers_id[loop],
                        "applicantId": self.xl_Applicant_id[loop],
                        "interviewerDecision": self.xl_decision[loop],
                        "interviewerComment": self.xl_Over_all_comment[loop],
                        "interviewDuration": self.xl_duration[loop],
                        "interviewedDate": self.xl_int_datetime[loop]
                    }]
                }
                providing_feedback = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/interview/givefeedback/",
                                                   headers=self.get_token,
                                                   data=json.dumps(feedback_request, default=str), verify=False)
                feedback_response = json.loads(providing_feedback.content)
                data = feedback_response.get('data')
                self.feedback_message = data.get('message')
                self.is_feedback = True
                print('Provide Feedback is Done')

                if self.xl_decision[loop] == 167097:
                    self.updatedecision = True

            except exceptions.ValueError as feedback_error:
                print(feedback_error)

    def feedback_details(self, loop):
        try:
            details_url = requests.get("https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}".format(self.ir),
                                       headers=self.get_token)
            details_response = json.loads(details_url.content)
            self.feedback_data = details_response['data']
            self.filledFeedbackDetails = self.feedback_data['filledFeedbackDetails']
            applicant = self.feedback_data['applicants']
            for applicants in applicant:
                self.applicant_details = applicants

            for feedback in self.filledFeedbackDetails:
                self.feedback = feedback
                for skillAssessed_details in feedback['skillAssessed']:
                    self.skillAssessed_details = skillAssessed_details

                    if self.xl_Skill_id_01[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_1 = skillAssessed_details

                    if self.xl_Skill_id_02[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_2 = skillAssessed_details

                    if self.xl_Skill_id_03[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_3 = skillAssessed_details

                    if self.xl_Skill_id_04[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_4 = skillAssessed_details

            print('Feedback details are fetched Successfully')
        except exceptions.ValueError as details_error:
            print(details_error)

    def updated_feedback_details(self, loop):
        try:
            details_url = requests.get("https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}".format(self.ir),
                                       headers=self.get_token)
            details_response = json.loads(details_url.content)
            self.updated_feedback_data = details_response['data']
            self.updated_filledFeedbackDetails = self.updated_feedback_data['filledFeedbackDetails']

            for updated_feedback in self.updated_filledFeedbackDetails:
                self.updated_feedback = updated_feedback
                for updated_skillAssessed_details in updated_feedback['skillAssessed']:
                    self.updated_skillAssessed_details = updated_skillAssessed_details

                    if self.xl_Skill_id_01[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_1 = updated_skillAssessed_details

                    if self.xl_Skill_id_02[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_2 = updated_skillAssessed_details

                    if self.xl_Skill_id_03[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_3 = updated_skillAssessed_details

                    if self.xl_Skill_id_04[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_4 = updated_skillAssessed_details

            print('Updated Feedback details are fetched Successfully')
        except exceptions.ValueError as details_error:
            print(details_error)

    def update_decision(self, loop):
        if self.xl_update_stage[loop]:
            update_decision_request = {
                "interviewRequestId": self.ir,
                "decisionId": self.xl_update_stage[loop]
            }
            update_decision_url = requests.post('https://amsin.hirepro.in/py/crpo/api/v1/interview/'
                                                'updateinterviewerdecision', headers=self.get_token,
                                                data=json.dumps(update_decision_request, default=str), verify=False)
            update_decision_response = json.loads(update_decision_url.content)
            decision_response = update_decision_response.get('data')
            decision_error = update_decision_response.get('error')
            self.decision_update = decision_response
            self.decision_error = decision_error

    def decision_updated_feedback_details(self, loop):
        try:
            details_url = requests.get("https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}".format(self.ir),
                                       headers=self.get_token)
            details_response = json.loads(details_url.content)
            decision_updated_feedback_data = details_response['data']
            decision_updated_filledfeedbackdetails = decision_updated_feedback_data['filledFeedbackDetails']

            for decision_updated_feedback in decision_updated_filledfeedbackdetails:
                self.decision_updated_feedback = decision_updated_feedback

        except exceptions.ValueError as decision:
            print(decision)

    def partial_feedback(self, loop):
        if self.feedback['partialFeedback'] == 1:

            try:
                update_feedback = {
                    "InterviewRequestId": self.ir,
                    "FilledFormId": self.skillAssessed_details['interviewfilledfeedbackformId'],
                    "Duration": self.xl_updated_duration[loop],
                    "OverAllComments": self.xl_Updated_Over_all_comment[loop],
                    "Skills": [{
                        "Id": self.skill_dict_1['id'],
                        "SkillId": self.xl_Skill_id_01[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_update_skill_score_01[loop]
                    }, {
                        "Id": self.skill_dict_2['id'],
                        "SkillId": self.xl_Skill_id_02[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_02[loop]
                    }, {
                        "Id": self.skill_dict_3['id'],
                        "SkillId": self.xl_Skill_id_03[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_03[loop]
                    }, {
                        "Id": self.skill_dict_4['id'],
                        "SkillId": self.xl_Skill_id_04[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_04[loop]
                    }]
                }
                partial_url = requests.post(
                    "https://amsin.hirepro.in/py/crpo/api/v1/interview/updateinterviewerfeedback",
                    headers=self.get_token,
                    data=json.dumps(update_feedback, default=str), verify=False)

                partial_response = json.loads(partial_url.content)
                self.partial_data = partial_response['data']

            except exceptions.ValueError as Partial_update_error:
                print(Partial_update_error)

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Expected', self.__style4)
        self.ws.write(self.rowsize+1, self.col, 'Actual', self.__style5)

        self.ws.write(self.rowsize, 1, self.xl_expected_status[loop], self.__style2)
        if self.applicant_details and self.applicant_details.get('applicantId'):
            if self.xl_expected_status[loop] == str('Pass') or str('Fail'):
                self.ws.write(self.rowsize+1, 1, 'Pass', self.__style2)
            else:
                self.ws.write(self.rowsize + 1, 1, 'Pass', self.__style3)
        else:
            if self.xl_expected_status[loop] == str('Fail'):
                self.ws.write(self.rowsize+1, 1, 'Fail', self.__style2)
            else:
                self.ws.write(self.rowsize + 1, 1, 'Fail', self.__style3)

        self.ws.write(self.rowsize, 2, self.xl_Applicant_id[loop], self.__style1)
        self.ws.write(self.rowsize + 1, 2, self.applicant_details.get('applicantId'), self.__style1)

        if self.is_success:
            self.ws.write(self.rowsize+1, 3, self.feedback_data['interviewerComment'], self.__style1)
        else:
            self.ws.write(self.rowsize+1, 3, self.message, self.__style1)

        if self.ir:
            self.ws.write(self.rowsize, 4, self.ir, self.__style1)
        else:
            self.ws.write(self.rowsize, 4, None)

        if self.is_feedback:
            self.ws.write(self.rowsize, 5, self.feedback_message, self.__style1)
        else:
            self.ws.write(self.rowsize, 5, None)

        if self.feedback and self.feedback['decisionText']:
            self.ws.write(self.rowsize, 6, self.feedback['decisionText'], self.__style1)
        else:
            self.ws.write(self.rowsize, 6, None)

        if self.feedback and self.feedback['partialFeedback'] == 1:
            self.ws.write(self.rowsize, 7, 'True', self.__style1)
        elif self.feedback and self.feedback['partialFeedback'] == 0:
            self.ws.write(self.rowsize, 7, 'False', self.__style1)

        if self.partial_data.get('message'):
            self.ws.write(self.rowsize, 8, self.partial_data['message'], self.__style1)
        else:
            self.ws.write(self.rowsize, 8, None)

        if self.decision_update and self.decision_update.get('message'):
            self.ws.write(self.rowsize, 9, self.decision_update.get('message'), self.__style1)
        elif self.decision_error and self.decision_error.get('errorDescription'):
            self.ws.write(self.rowsize, 9, self.decision_error.get('errorDescription'), self.__style1)
        else:
            self.ws.write(self.rowsize, 9, None)

        if self.decision_updated_feedback and self.decision_updated_feedback['decisionText']:
            self.ws.write(self.rowsize, 10, self.decision_updated_feedback['decisionText'], self.__style1)
        else:
            self.ws.write(self.rowsize, 10, None)

        self.ws.write(self.rowsize, 11, self.xl_Datetime[loop], self.__style1)
        if self.xl_Datetime[loop] == self.feedback_data.get('interviewTime'):
            if self.xl_Datetime[loop] is None:
                self.ws.write(self.rowsize+1, 11, 'Empty_value', self.__style6)
            else:
                self.ws.write(self.rowsize+1, 11, self.feedback_data.get('interviewTime'))
        else:
            self.ws.write(self.rowsize+1, 11, self.feedback_data.get('interviewTime'))

        self.ws.write(self.rowsize, 12, self.xl_int_datetime[loop], self.__style1)
        if self.is_feedback:
            if self.xl_int_datetime[loop] == self.feedback.get('interviewedTime'):
                if self.xl_int_datetime[loop] is None:
                    self.ws.write(self.rowsize+1, 12, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 12, self.feedback.get('interviewedTime'))
            else:
                self.ws.write(self.rowsize+1, 12, self.feedback.get('interviewedTime'))
        else:
            self.ws.write(self.rowsize+1, 12, None)

        self.ws.write(self.rowsize, 13, self.xl_Skill_id_01[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_id_01[loop] == self.skill_dict_1.get('skillId'):
                if self.xl_Skill_id_01[loop] is None:
                    self.ws.write(self.rowsize + 1, 13, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 13, self.skill_dict_1.get('skillId'))
        else:
            self.ws.write(self.rowsize + 1, 13, None)

        self.ws.write(self.rowsize, 14, self.xl_Skill_score_01[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_score_01[loop] == self.skill_dict_1.get('skillScore'):
                if self.xl_Skill_score_01[loop] is None:
                    self.ws.write(self.rowsize + 1, 14, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 14, self.skill_dict_1.get('skillScore'))
        else:
            self.ws.write(self.rowsize + 1, 14, None)

        self.ws.write(self.rowsize, 15, self.xl_Skill_id_02[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_id_02[loop] == self.skill_dict_2.get('skillId'):
                if self.xl_Skill_id_02[loop] is None:
                    self.ws.write(self.rowsize+1, 15, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 15, self.skill_dict_2.get('skillId'))
        else:
            self.ws.write(self.rowsize+1, 15, None)

        self.ws.write(self.rowsize, 16, self.xl_Skill_score_02[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_score_02[loop] == self.skill_dict_2.get('skillScore'):
                if self.xl_Skill_score_02[loop] is None:
                    self.ws.write(self.rowsize+1, 16, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 16, self.skill_dict_2.get('skillScore'))
        else:
            self.ws.write(self.rowsize+1, 16, None)

        self.ws.write(self.rowsize, 17, self.xl_Skill_id_03[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_id_03[loop] == self.skill_dict_3.get('skillId'):
                if self.xl_Skill_id_03[loop] is None:
                    self.ws.write(self.rowsize + 1, 17, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 17, self.skill_dict_3.get('skillId'))
        else:
            self.ws.write(self.rowsize + 1, 17, None)

        self.ws.write(self.rowsize, 18, self.xl_Skill_score_03[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_score_03[loop] == self.skill_dict_3.get('skillScore'):
                if self.xl_Skill_score_03[loop] is None:
                    self.ws.write(self.rowsize+1, 18, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 18, self.skill_dict_3.get('skillScore'))
        else:
            self.ws.write(self.rowsize+1, 18, None)

        self.ws.write(self.rowsize, 19, self.xl_Skill_id_04[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_id_04[loop] == self.skill_dict_4.get('skillId'):
                if self.xl_Skill_id_04[loop] is None:
                    self.ws.write(self.rowsize + 1, 19, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 19, self.skill_dict_4.get('skillId'))
        else:
            self.ws.write(self.rowsize + 1, 19, None)

        self.ws.write(self.rowsize, 20, self.xl_Skill_score_04[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Skill_score_04[loop] == self.skill_dict_4.get('skillScore'):
                if self.xl_Skill_score_04[loop] is None:
                    self.ws.write(self.rowsize + 1, 20, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 20, self.skill_dict_4.get('skillScore'))
        else:
            self.ws.write(self.rowsize + 1, 20, None)

        self.ws.write(self.rowsize, 21, self.xl_duration[loop], self.__style1)
        if self.is_feedback:
            if self.xl_duration[loop] == self.feedback.get('duration'):
                if self.xl_duration[loop] is None:
                    self.ws.write(self.rowsize + 1, 21, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 21, self.feedback.get('duration'))
        else:
            self.ws.write(self.rowsize + 1, 21, None)

        self.ws.write(self.rowsize, 22, self.xl_skill_comment[loop], self.__style1)
        if self.is_feedback:
            if self.xl_skill_comment[loop] == self.skill_dict_4.get('skillComment'):
                if self.xl_Skill_score_04[loop] is None:
                    self.ws.write(self.rowsize+1, 22, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 22, self.skill_dict_4.get('skillComment'))
        else:
            self.ws.write(self.rowsize+1, 22, None)

        self.ws.write(self.rowsize, 23, self.xl_Over_all_comment[loop], self.__style1)
        if self.is_feedback:
            if self.xl_Over_all_comment[loop] == self.feedback.get('comment'):
                if self.xl_Over_all_comment[loop] is None:
                    self.ws.write(self.rowsize + 1, 23, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize + 1, 23, self.feedback.get('comment'))
        else:
            self.ws.write(self.rowsize + 1, 23, None)

        self.ws.write(self.rowsize, 24, self.xl_update_skill_score_01[loop], self.__style1)
        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_update_skill_score_01[loop] == self.updated_skill_dict_1.get('skillScore'):
                if self.xl_update_skill_score_01[loop] is None:
                    self.ws.write(self.rowsize+1, 24, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 24, self.updated_skill_dict_1.get('skillScore'))
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 24, self.updated_skill_dict_1.get('skillScore', 'No_Update_Details'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 24, self.updated_skill_dict_1.get('skillScore', 'NA'), self.__style1)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 24,
                                  self.updated_skill_dict_1.get('skillScore', 'Not a Partial/Update feedback'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 24, self.updated_skill_dict_1.get('skillScore', 'NA'), self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 25, self.xl_updated_duration[loop], self.__style1)
        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_updated_duration[loop] == self.updated_feedback.get('duration'):
                if self.xl_updated_duration[loop] is None:
                    self.ws.write(self.rowsize+1, 25, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 25, self.updated_feedback.get('duration'))
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 25, self.updated_feedback.get('duration', 'No_Update_Details'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 25, self.updated_feedback.get('duration', 'NA'), self.__style1)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 25,
                                  self.updated_feedback.get('duration', 'Not a Partial/Update feedback'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 25, self.updated_feedback.get('duration', 'NA'), self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 26, self.xl_update_Skill_comment[loop], self.__style1)
        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_update_Skill_comment[loop] == self.updated_skill_dict_4.get('skillComment'):
                if self.xl_update_Skill_comment[loop] is None:
                    self.ws.write(self.rowsize+1, 26, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 26, self.updated_skill_dict_4.get('skillComment'))
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 26, self.updated_skill_dict_4.get('skillComment', 'No_Update_Details'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 26, self.updated_skill_dict_4.get('skillComment', 'NA'), self.__style1)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 26,
                                  self.updated_skill_dict_4.get('skillComment', 'Not a Partial/Update feedback'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 26, self.updated_skill_dict_4.get('skillComment', 'NA'), self.__style1)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 27, self.xl_Updated_Over_all_comment[loop], self.__style1)
        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_Updated_Over_all_comment[loop] == self.updated_feedback.get('comment'):
                if self.xl_Updated_Over_all_comment[loop] is None:
                    self.ws.write(self.rowsize+1, 27, 'Empty_value', self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 27, self.updated_feedback.get('comment'))
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 27, self.updated_feedback.get('comment', 'No_Update_Details'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 27, self.updated_feedback.get('comment', 'NA'), self.__style1)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize+1, 27,
                                  self.updated_feedback.get('comment', 'Not a Partial/Update feedback'),
                                  self.__style6)
                else:
                    self.ws.write(self.rowsize+1, 27, self.updated_feedback.get('comment', 'NA'), self.__style1)

        self.rowsize += 2
        Object.wb_Result.save('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/'
                              'CRPO/Give Feedback/API_GiveFeedback.xls')


Object = InterviewFeedback()
Object.excel_data()
Total_count = len(Object.xl_Event_id)
print "Number of Rows ::", Total_count

try:
    if Object.login == 'OK':
        for looping in range(0, Total_count):
            print "Iteration Count is ::", looping
            Object.schedule_interview(looping)
            if Object.is_success:
                Object.provide_feedback(looping)
                Object.feedback_details(looping)
                if Object.pf:
                    Object.partial_feedback(looping)
                    Object.updated_feedback_details(looping)
                if Object.updatedecision:
                    Object.update_decision(looping)
                    Object.decision_updated_feedback_details(looping)
            Object.output_excel(looping)
            print('Excel data is ready')

            # -------------------------------------
            # Making all dict empty for every loop
            # -------------------------------------
            Object.is_success = {}
            Object.message = {}
            Object.ir = {}
            Object.feedback_data = {}
            Object.feedback = {}
            Object.is_feedback = {}

            Object.updated_feedback_data = {}
            Object.updated_feedback = {}
            Object.partial_data = {}
            Object.feedback_message = {}
            Object.applicant_details = {}

            # ----------
            # Skill dict
            # ----------
            Object.skill_dict_1 = {}
            Object.skill_dict_2 = {}
            Object.skill_dict_3 = {}
            Object.skill_dict_4 = {}

            Object.filledFeedbackDetails = {}
            Object.skillAssessed_details = {}

            # ------------------
            # updated Skill dict
            # ------------------
            Object.updated_skill_dict_1 = {}
            Object.updated_skill_dict_2 = {}
            Object.updated_skill_dict_3 = {}
            Object.updated_skill_dict_4 = {}

            Object.updated_filledFeedbackDetails = {}
            Object.updated_skillAssessed_details = {}

            Object.decision_update = {}
            Object.updatedecision = {}
            Object.decision_error = {}
            Object.decision_updated_feedback = {}
            # ---------------------
            # Partial/updated  dict
            # ---------------------
            Object.pf = {}

except exceptions.AttributeError as Object_error:
    print(Object_error)
