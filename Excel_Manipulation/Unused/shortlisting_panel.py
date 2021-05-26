from Excel_Manipulation.read_excel import *
import mysql.connector
import itertools
from operator import itemgetter
import requests
import json
import datetime
import xlwt
import time


class SCAutomation:
    def __init__(self):
        now = datetime.datetime.now()
        # self.LoginName = raw_input('LoginName:: ')
        # self.Password = raw_input('Password:: ')
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.rowsize = 1
        file_path = '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_InputData/' \
                    'CRPO/ShortlistingPanel/Test.xls'
        sheet_index = 0
        excel_read_obj.excel_read(file_path, sheet_index)

        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.__style4 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                    'font: name Times New Roman, color-index black, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('SLC_Verification')
        self.ws.write(0, 0, 'CandidateID', self.__style4)
        self.ws.write(0, 1, 'ApplicantID', self.__style4)
        self.ws.write(0, 2, 'EventID', self.__style4)
        self.ws.write(0, 3, 'JobID', self.__style4)
        self.ws.write(0, 4, 'SLCID', self.__style4)
        self.ws.write(0, 5, 'TestID', self.__style4)
        self.ws.write(0, 6, 'Excel Expected Status', self.__style4)
        self.ws.write(0, 7, 'DB Status', self.__style4)
        self.ws.write(0, 8, 'Excel - DB Match Result', self.__style4)

        # CRPO LOGIN APPLICATION
        self.header = {"content-type": "application/json"}
        self.data = {"LoginName": 'admin',
                     "Password": '4LWS-0671',
                     "TenantAlias": "Automation",
                     "UserName": 'admin'}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=self.header,
                                 data=json.dumps(self.data), verify=False)
        self.abc = response.json()
        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.abc.get("Token")}
        print self.headers

    # Making AMS DB Connection
    def amsConnection(self):
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()

    # Reading XL data
    def applicantDataRead(self):
        # self.kk1 = []
        self.xl_applicant_id = []
        self.applicant_json_data = []
        for i in excel_read_obj.details:
            self.xl_applicant_id.append(int(i.get('applicantId')))
            # print self.xl_applicant_id
            local = str(i.get('testId'))
            applicant_test_id = [int(float(b)) for b in local.split(',')]

            local = str(i.get('scId'))
            applicant_scid = [int(float(b)) for b in local.split(',')]

            self.convert_json = {"applicantId": int(i.get('applicantId')), "eventId": int(i.get('eventId')),
                                 "jobId": int(i.get('jobId')), "mjrId": int(i.get('mjrId')),
                                 "mjrStatusId": int(i.get('mjrStatusId')), "testId": applicant_test_id,
                                 "scId": applicant_scid}
            self.applicant_json_data.append(self.convert_json)
            # print self.applicant_json_data
            # print type(self.applicant_json_data)
        self.totalapplicantCount = len(self.xl_applicant_id)
        a = self.xl_applicant_id
        self.xl_all_applicant_id = ','.join(str(v) for v in a)
        print self.xl_all_applicant_id
        print self.totalapplicantCount

    def groupby_MJR_TEST_SLC(self):

        # Sort applicant data by `mjrid, Testid,scid,JobId and Eventid` key.
        self.applicant_json_data = sorted(self.applicant_json_data,
                                          key=itemgetter('eventId', 'jobId', 'mjrId', 'testId', 'scId'))
        # print self.applicant_json_data
        # Display data grouped by `mjrid, Testid,scid,JobId and Eventid` key.
        for key, value in itertools.groupby(self.applicant_json_data,
                                            key=itemgetter('eventId', 'jobId', 'mjrId', 'testId', 'scId')):
            # print key
            self.all_mjr_applicants = []
            self.to_status_id = []
            for iter in self.applicant_json_data:
                self.value = iter
                if key[0] == self.value['eventId'] and key[1] == self.value['jobId'] and \
                        key[2] == self.value['mjrId'] and key[3] == self.value['testId'] and key[4] == self.value[
                    'scId']:
                    self.all_mjr_applicants.append(self.value['applicantId'])
                    self.to_status_id.append(self.value['mjrStatusId'])
            # print self.all_mjr_applicants
            # print self.to_status_id
            self.data = {"ApplicantIds": self.all_mjr_applicants,
                         "EventId": key[0],
                         "JobRoleId": key[1],
                         "ToStatusId": self.to_status_id[0],
                         "Sync": "True", "Comments": "",
                         "InitiateStaffing": False,
                         "MjrId": key[2]}
            print  self.data
            time.sleep(3)
            r = requests.post("https://amsin.hirepro.in/py/crpo/applicant/api/v1/applicantStatusChange/",
                              headers=self.headers, data=json.dumps(self.data, default=str), verify=False)
            resp_dict = json.loads(r.content)
            self.status = resp_dict['status']
            if self.status == 'OK':
                print "API Status is", self.status

                self.data = {"eventId": key[0],
                             "jobId": key[1],
                             "statusId": self.to_status_id[0],
                             "testIds": key[3],
                             "shortlistingCriteriaIds": key[4]}
                print self.data
                # time.sleep(3)
                r = requests.post("https://amsin.hirepro.in/py/crpo/shortlistingcriteria/api/v1/oneClickShortlist",
                                  headers=self.headers, data=json.dumps(self.data, default=str), verify=False)
                resp_dict = json.loads(r.content)
                # print resp_dict

    # pending is get all applicant statuss, match excel, writeexcel
    def allApplicantStatuss(self):
        try:
            self.amsConnection()
            self.applicant_query = "select a.id,rs.label as currentstatus from applicant_statuss a " \
                                   "left join resume_statuss rs on a.current_status_id=rs.id where a.id " \
                                   "in(%s) order by id asc;" % (self.xl_all_applicant_id)

            query = self.applicant_query
            time.sleep(2)
            self.cursor.execute(query)
            db_applicant_details = self.cursor.fetchall()
            self.db_all_applicants = []
            self.db_applicant_id = []
            self.db_applicant_status = []
            self.db_j = 0

            # converting json is important for comparision
            for new_data in db_applicant_details:
                self.db_applicant_id.append(new_data[0])
                self.db_applicant_status.append(new_data[1])
                self.convert_json1 = {'dbApplicantID': self.db_applicant_id[self.db_j],
                                      'dbApplicantStatus': self.db_applicant_status[self.db_j]}
                self.db_j += 1
                self.db_all_applicants.append(self.convert_json1)
        except:
            print("DB connection Error")

    def match_db_excel(self):
        total_db_count = len(self.db_all_applicants)
        for iteration_count in range(0, self.totalapplicantCount):
            for dbdata in range(0, total_db_count):
                dbvalue = self.db_all_applicants[dbdata]
                excel_value = excel_read_obj.details[iteration_count]
                if excel_value.get('applicantId') == dbvalue['dbApplicantID']:
                    if excel_value.get('expectedStatus') == dbvalue['dbApplicantStatus']:
                        self.db_mess = 'pass'
                        self.excel_write(excel_value.get('candidateID'),
                                         excel_value.get('applicantId'),
                                         excel_value.get('expectedStatus'),
                                         dbvalue['dbApplicantStatus'],
                                         self.db_mess,
                                         self.__style3, excel_value.get('eventId'),
                                         excel_value.get('jobId'),
                                         excel_value.get('testId'),
                                         excel_value.get('scId'))
                    else:
                        self.db_mess = 'status not matched with excel and DB'
                        self.excel_write(excel_value.get('candidateID'),
                                         excel_value.get('applicantId'),
                                         excel_value.get('expectedStatus'),
                                         dbvalue['dbApplicantStatus'],
                                         self.db_mess,
                                         self.__style3, excel_value.get('eventId'),
                                         excel_value.get('jobId'),
                                         excel_value.get('testId'),
                                         excel_value.get('scId'))

    def excel_write(self, candidate_id, xl_aid, xl_exp_status, db_status, db_mess, __style, event_id, job_id, testid,
                    scid):
        self.ws.write(self.rowsize, 0, candidate_id, self.__style0)
        self.ws.write(self.rowsize, 1, xl_aid, self.__style0)
        self.ws.write(self.rowsize, 2, event_id, self.__style0)
        self.ws.write(self.rowsize, 3, job_id, self.__style0)
        self.ws.write(self.rowsize, 4, str(testid), self.__style0)
        self.ws.write(self.rowsize, 5, str(scid), self.__style0)
        self.ws.write(self.rowsize, 6, xl_exp_status, __style)
        self.ws.write(self.rowsize, 7, db_status, __style)
        self.ws.write(self.rowsize, 8, db_mess, __style)
        self.rowsize = self.rowsize + 1
        ob.wb_Result.save('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/CRPO/CRPO_SLC'
                          '/API_Shortlisting_Panel(' + self.__current_DateTime + ').xls')


ob = SCAutomation()
ob.applicantDataRead()
ob.groupby_MJR_TEST_SLC()
time.sleep(5)
ob.allApplicantStatuss()
ob.match_db_excel()
