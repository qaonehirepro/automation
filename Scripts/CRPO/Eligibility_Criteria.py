import mysql.connector
import requests
import json
import xlrd
import datetime
import xlwt
import time


class ECAutomation:
    def __init__(self):
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.rowsize = 1
        # excel values are assigned to local variable
        self.xl_candidate_id = []
        self.xl_applicant_id = []
        self.xl_expected_result = []
        self.xl_ec_id = []
        self.xl_event_Id = []
        self.xl_job_Id = []
        self.xl_job_Id = []
        self.xl_mjr_Id = []
        self.xl_to_status__Id = []
        self.xl_positive_status__Id = []
        self.xl_negative_status__Id = []
        self.xl_ec_configuration__Id = []
        self.__style4 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                    'font: name Times New Roman, color-index black, bold on')
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('EC_Verification')
        self.ws.write(0, 0, 'Event Id', self.__style4)
        self.ws.write(0, 1, 'Job Id', self.__style4)
        self.ws.write(0, 2, 'CandidateID', self.__style4)
        self.ws.write(0, 3, 'Excel ApplicantId', self.__style4)
        self.ws.write(0, 4, 'DB Applicant ID', self.__style4)
        self.ws.write(0, 5, 'Excel Expected Status', self.__style4)
        self.ws.write(0, 6, 'DB Status', self.__style4)
        self.ws.write(0, 7, 'EC ID', self.__style4)
        self.ws.write(0, 8, 'Excel - DB Match Result', self.__style4)
        # CRPO LOGIN APPLICATION
        self.header = {"content-type": "application/json"}
        self.data = {"LoginName": 'admin',
                     "Password": '4LWS-067',
                     "TenantAlias": "Automation",
                     "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/",
                                 headers=self.header,
                                 data=json.dumps(self.data),
                                 verify=False)
        self.abc = response.json()
        self.headers = {"content-type": "application/json",
                        "X-AUTH-TOKEN": self.abc.get("Token")}
        self.var = None
        time.sleep(1)
        resp_dict = json.loads(response.content)
        self.status = resp_dict['status']
        if self.status == 'OK':
            self.var = 'OK'
            print "Login successfully"
            print "Status is", self.status
            time.sleep(1)
        else:
            self.var = 'KO'
            print "Failed to login"
            print "Status is", self.status

    # -------------------------------------------
    # Reading Input data from excel
    # -------------------------------------------
    def Data_read(self):
        wb = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/'
                                'PythonWorkingScripts_InputData/CRPO/EC/Final_EC_Configuration.xls')
        sh1 = wb.sheet_by_index(0)
        i = 1
        for i in range(1, sh1.nrows):
            rownum = i
            rows = sh1.row_values(rownum)
            self.xl_candidate_id.append(int(rows[0]))
            self.xl_applicant_id.append(int(rows[1]))
            self.xl_expected_result.append(rows[2])
            self.xl_ec_id.append(int(rows[3]))
            self.xl_event_Id.append(int(rows[4]))
            self.xl_job_Id.append(int(rows[5]))
            self.xl_mjr_Id.append(int(rows[6]))
            self.xl_to_status__Id.append(int(rows[7]))
            self.xl_positive_status__Id.append(int(rows[8]))
            self.xl_negative_status__Id.append(int(rows[9]))
            self.xl_ec_configuration__Id.append(int(rows[10]))
            self.applicant_id = self.xl_applicant_id
            self.candidate_id = self.xl_candidate_id
            self.expected_results = self.xl_expected_result
            self.ec_id = self.xl_ec_id

    # ------------------------------------------------------------------------------------------------------------------
    # Fetching applicant status data from DB
    # below method is used for Database connectivity and query execution
    # conn.close is important for every iteration otherwise, it will wrong data from local image.
    # ------------------------------------------------------------------------------------------------------------------
    def Fetch_Applicants_DB(self, iteration_count):
        try:
            conn = mysql.connector.connect(host='35.154.36.218',
                                           database='appserver_core',
                                           user='hireprouser',
                                           password='tech@123')
            cursor = conn.cursor()
            self.applicant_query = "select  ap.candidate_id,ap.id as applicant_id,ap.current_status_id,rs.label as status " \
                                   "from applicant_statuss ap left join resume_statuss rs on ap.current_status_id = rs.id " \
                                   "where ap.id ='%s';" % (self.applicant_id[iteration_count])
            query = self.applicant_query
            time.sleep(2)
            cursor.execute(query)
            data_to_return = cursor.fetchall()
            conn.close()
            alldata = data_to_return
            tot = len(alldata)
            if tot > 0:
                for item in range(0, tot):
                    self.data1 = alldata[item]
                    self.db_candidate_id = self.data1[0]
                    self.db_applicant_id = self.data1[1]
                    self.db_status_id = self.data1[2]
                    self.db_status = self.data1[3]
            else:
                print "No Data From Database"
        except:
            print("DB connection Error")

    def api_main(self, iteration_count):
        # -------------------------------------------
        # Updating EC configuration at event level
        # -------------------------------------------
        self.data = {"ecConfigurations": [{"id": self.xl_ec_configuration__Id[iteration_count],
                                           "jobRoleId": self.xl_job_Id[iteration_count],
                                           "eventId": self.xl_event_Id[iteration_count],
                                           "ecId": self.ec_id[iteration_count],
                                           "positiveStatusId": self.xl_positive_status__Id[iteration_count],
                                           "negativeStatusId": self.xl_negative_status__Id[iteration_count]}]}
        r = requests.post("https://amsin.hirepro.in/py/crpo/dynamicec/api/v1/createOrUpdateEcConfig/",
                          headers=self.headers,
                          data=json.dumps(self.data, default=str),
                          verify=False)
        time.sleep(1)
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        if self.status == 'OK':
            print "EC Updated Successfully"
            print "EC Status Code is", self.status
            # ---------------------------
            # Changing applicant status
            # ---------------------------
            self.data = {"ApplicantIds": [self.applicant_id[iteration_count]],
                         "EventId": self.xl_event_Id[iteration_count],
                         "JobRoleId": self.xl_job_Id[iteration_count],
                         "ToStatusId": self.xl_to_status__Id[iteration_count],
                         "Sync": "True",
                         "Comments": "",
                         "InitiateStaffing": False,
                         "MjrId": self.xl_mjr_Id[iteration_count]}
            r = requests.post("https://amsin.hirepro.in/py/crpo/applicant/api/v1/applicantStatusChange/",
                              headers=self.headers,
                              data=json.dumps(self.data, default=str), verify=False)
            resp_dict = json.loads(r.content)
            self.status = resp_dict['status']
            if self.status == 'OK':
                print "Status Change API is Executed"
                print "Status is", self.status
                time.sleep(2)
            else:
                print "Status Change API is not executed"
                print "Status is", self.status
        else:
            print "EC Not Updated Successfully"
            print "EC Status Code is", self.status

    # -------------------------------------
    # Compairing data from Excel to DB
    # -------------------------------------
    def match_db_excel(self, iteration_count):
        if self.applicant_id[iteration_count] == self.db_applicant_id:
            if self.expected_results[iteration_count] == self.db_status:
                print "DB - Status Matched With Expected Status"
                self.db_mess = 'pass'
                self.excel_write(self.candidate_id[iteration_count],
                                 self.applicant_id[iteration_count],
                                 self.db_applicant_id,
                                 self.expected_results[iteration_count],
                                 self.db_status,
                                 self.db_mess,
                                 self.__style3,
                                 self.xl_event_Id[iteration_count],
                                 self.xl_job_Id[iteration_count],
                                 self.ec_id[iteration_count])
            else:
                print "DB - Status Not Matched  With Expected Status"
                self.db_mess = 'status not matched with excel and DB'
                self.excel_write(self.candidate_id[iteration_count],
                                 self.applicant_id[iteration_count],
                                 self.db_applicant_id,
                                 self.expected_results[iteration_count],
                                 self.db_status,
                                 self.db_mess,
                                 self.__style2,
                                 self.xl_event_Id[iteration_count],
                                 self.xl_job_Id[iteration_count],
                                 self.ec_id[iteration_count])
        else:
            pass
            self.db_mess = 'Excel Applicant id not matched with DB applicant'
            self.excel_write(self.candidate_id[iteration_count],
                             self.applicant_id[iteration_count],
                             self.db_applicant_id,
                             self.expected_results[iteration_count],
                             self.db_status,
                             self.db_mess,
                             self.__style2,
                             self.xl_event_Id[iteration_count],
                             self.xl_job_Id[iteration_count],
                             self.ec_id[iteration_count])

    # ---------------------------
    # Writing output to excel
    # ---------------------------
    def excel_write(self, candidate_id, ex_aid, db_aid, ex_exp_status, db_status, db_mess, __style, event_id, job_id,
                    ec_id):
        self.ws.write(self.rowsize, 0, event_id, self.__style0)
        self.ws.write(self.rowsize, 1, job_id, self.__style0)
        self.ws.write(self.rowsize, 2, candidate_id, self.__style0)
        self.ws.write(self.rowsize, 3, ex_aid, self.__style0)
        self.ws.write(self.rowsize, 4, db_aid, self.__style0)
        self.ws.write(self.rowsize, 5, ex_exp_status, self.__style0)
        self.ws.write(self.rowsize, 6, db_status, self.__style0)
        self.ws.write(self.rowsize, 7, ec_id, self.__style0)
        self.ws.write(self.rowsize, 8, db_mess, __style)
        self.rowsize = self.rowsize + 1
        ob.wb_Result.save(
            '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/CRPO/EC'
            '/API_DB_EC_Verification(' + self.__current_DateTime + ').xls')
    # Class Invocation


ob = ECAutomation()
ob.Data_read()
tot_count = len(ob.applicant_id)
print tot_count
if ob.var == 'OK':
    for iteration_count in range(0, tot_count):
        print "Iteration Count is '%d'" % iteration_count
        ob.api_main(iteration_count)
        ob.Fetch_Applicants_DB(iteration_count)
        ob.match_db_excel(iteration_count)
