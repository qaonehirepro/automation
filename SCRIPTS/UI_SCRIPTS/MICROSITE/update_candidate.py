from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.io_path import *
from selenium import webdriver
import time
import datetime
import xlwt
import mysql
import mysql.connector
import requests
import json

class CreateCase:

    def __init__(self):

        sprint_id = input('Enter Sprint ID')
        self.create_candidate_request = {"PersonalDetails": {"Name": "Muthu Murugan Ramalingam",
                                                             "Email1": 'S1N1J1E1V1' + sprint_id + '@gmail.com',
                                                             "USN": sprint_id},
                                         "SourceDetails": {"SourceId": "2366"}}
        self.create_candidate()
        self.starttime = datetime.datetime.now()
        self.starttime1 = "Strated:- %s" % self.starttime.strftime(" %H-%M-%S")


        self.current_DateTime = self.starttime.strftime("%d-%m-%Y")
        self.style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.over_all_status = 'Pass'
        self.over_all_status_color = xlwt.easyxf(
            'font: bold on,height 250,color-index green;pattern: pattern solid,fore-colour light_yellow;'
            'border: left thin,right thin,top thin,bottom thin')

        self.base_style1 = xlwt.easyxf(
            'font: bold on,height 250,color-index black;pattern: pattern solid,fore-colour light_yellow;'
            'border: left thin,right thin,top thin,bottom thin')

        self.base_style2 = xlwt.easyxf(
            'font: bold on,height 250,color-index red;pattern: pattern solid,fore-colour light_yellow;'
            'border: left thin,right thin,top thin,bottom thin')

        # file_path = 'F:\\automation\\PythonWorkingScripts_InputData\\' \
        #             'Microsite\\GenericExcelTest.xls'
        file_path = input_path_microsite_update_case

        sheet_index = 1
        excel_read_obj.excel_read(file_path, sheet_index)

    def loginToCRPO(self):
        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "4LWS-0671", "TenantAlias": "Automation", "UserName": "admin"}
        response = requests.post('https://amsin.hirepro.in/py/common/user/login_user/', headers=header,
                                 data=json.dumps(data), verify=False)
        abc = response.json()
        headers = {"content-type": "application/json", "X-AUTH-TOKEN": abc.get("Token")}
        return headers

    def create_candidate(self):
        response = requests.post('https://amsin.hirepro.in/py/rpo/create_candidate/', headers=self.loginToCRPO(),
                                 data=json.dumps(self.create_candidate_request), verify=False)
        self.response_data = response.json()
        self.candidate_id = self.response_data.get('CandidateId')
        if self.response_data.get('status') == 'OK':
            print ("candidate created in crpo")
            # self.browser_location = "/home/muthumurugan/Desktop/chromedriver_2.37"
            self.url = 'https://automation-in.hirepro.in/?candidate=%s' % self.candidate_id
            # self.driver = webdriver.Chrome(self.browser_location)
            self.driver = webdriver.Chrome(executable_path=r"F:\qa_automation\chromedriver.exe")
        else:
            print ("candidate not created in CRPO due to some technical glitch")
            print (self.response_data)

    def msdbconnection(self):

        self.conn1 = mysql.connector.connect(host='52.66.103.86',
                                             database='automation-in052018',
                                             user='root',
                                             password='data')
        self.cursor = self.conn1.cursor()

    def amsdbconnection(self):
        # 35.154.213.175
        # 35.154.36.218
        self.conn = mysql.connector.connect(host='35.154.213.175',
                                            database='appserver_core',
                                            user='qauser',
                                            password='qauser')
        self.cursor = self.conn.cursor()

    def find_element_by_id(self, locator_id, xl_data, message):
        try:
            if self.driver.find_element_by_id(locator_id).is_displayed() != 0:
                self.data = xl_data
                find_element = self.driver.find_element_by_id(locator_id)
                find_element.clear()
                find_element.send_keys(self.data)

        except Exception as e:
            print (message)
            self.data = None
        return self.data

    def find_element_by_xpath_truefalse(self, locator_id, xl_data, message):
        try:
            if self.driver.find_element_by_xpath(locator_id).is_displayed() != 0:
                self.data = xl_data
                find_element = self.driver.find_element_by_xpath(locator_id).click()

        except Exception as e:
            print (message)
            self.data = None
        return self.data

    def find_element_by_xpath(self, locator_id, xl_data, message):
        try:
            time.sleep(0.3)
            locator_id1 = "//select[@id='%s']" % locator_id
            if self.driver.find_element_by_xpath(locator_id1).is_displayed() == 0:
                self.data = xl_data
                input = "//select[@id='%s']/option[text()='%s']" % (locator_id, self.data)
                find_element = self.driver.find_element_by_xpath(input).click()

        except Exception as e:
            print (message)
            self.data = None
        return self.data

    def declaration(self):
        try:
            if self.driver.find_element_by_id('declaration').is_displayed() != 0:
                self.driver.find_element_by_id('declaration').click()
                time.sleep(0.10)
        except Exception as e:
            print ("declaration is not available")

    def submit(self):
        try:
            if self.driver.find_element_by_xpath(".//*[@id='registerbtndiv']").is_displayed() != 0:
                self.driver.find_element_by_xpath(".//*[@id='registerbtndiv']").click()
                time.sleep(0.10)
        except:
            print ("submit button is not available")

    def microsite_query(self):

        C2.msdbconnection()
        self.ms_data = {}
        self.Candidate_Id = "select candidate_id,candidate_amsid from candidate where candidate_email = '%s' " \
                            "order by candidate_id desc limit 1" % (self.ui_email)
        query = self.Candidate_Id
        print (query)
        C2.execute_query(query)
        if self.data != None:
            self.Candidate_MS_Id = self.data[0]
            self.candidate_amsid = self.data[1]
            print (self.Candidate_MS_Id)
            print (self.candidate_amsid)
            self.MS_Cand_Personal = "select c.candidate_fname,c.candidate_mname,c.candidate_lname,c.candidate_email," \
                                    "c.candidate_email2,c.candidate_code,c.candidate_mobile,c.alternate_mobile," \
                                    "c.passport_no,c.pan_number,c.candidate_marital,cv.catalog_value_name as nationality," \
                                    "c.candidate_gender,aadhaarno,currentctc from candidate c left join catalog_values cv on " \
                                    "c.candidate_nationality= cv.ams_id  where c.candidate_id= %s;" \
                                    % (self.Candidate_MS_Id)
            query = self.MS_Cand_Personal
            C2.execute_query(query)
            self.ms_data.update({'firstName': self.data[0], 'middleName': self.data[1],
                                 'lastName': self.data[2], 'email': self.data[3],
                                 'altEmail': self.data[4],
                                 'usn': self.data[5], 'mobile': self.data[6],
                                 'altMobile': self.data[7],
                                 'passport': self.data[8], 'pancard': self.data[9],
                                 'maritalStatus': self.data[10],
                                 'nationality': self.data[11], 'gender': self.data[12],
                                 'aadhar': self.data[13],
                                 'currentCtc': self.data[14]})

            self.MS_Cand_Address = "select c.current_address1,c.current_address2," \
                                   "v2.catalog_value_name as Current_City,v.catalog_value_name as Current_state," \
                                   "c.current_pincode,c.permanent_address1,c.permanent_address2," \
                                   "v3.catalog_value_name as Permanent_City," \
                                   "v1.catalog_value_name  as Permanent_state,c.permanent_pincode, " \
                                   "v4.catalog_value_name as country " \
                                   "from candidate c inner join catalog_values v on c.current_state =v.ams_id " \
                                   "inner join catalog_values v1 on c.permanent_state =v1.ams_id " \
                                   "inner join catalog_values v2 on c.current_city =v2.ams_id " \
                                   "inner join catalog_values v3 on c.permanent_city =v3.ams_id " \
                                   "inner join catalog_values v4 on c.country=v4.ams_id " \
                                   "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Address
            C2.execute_query(query)
            self.ms_data.update(
                {'currentAddress1': self.data[0], 'currentAddress2': self.data[1], 'currentCity': self.data[2],
                 'currentState': self.data[3], 'currentPincode': self.data[4],
                 'permanentAddress1': self.data[5],
                 'permanentAddress2': self.data[6], 'permanentCity': self.data[7], 'permanentState': self.data[8],
                 'permanentPincode': self.data[9], 'country': self.data[10]})

            self.MS_Cand_Work_Experience = "select current_employer,current_experience,bpo_experience," \
                                           "total_experience,experience_in_year " \
                                           "from candidate where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Work_Experience
            C2.execute_query(query)
            self.ms_data.update({'currentCompany': self.data[0], 'currentExperience': self.data[1],
                                 'bpoExperience': self.data[2], 'totalExperience': self.data[3],
                                 'experienceInYear': self.data[4]})

            self.MS_Cand_Custom_text = "select text1,text2,text3,text4,text5,text6,text7,text8,text9,text10," \
                                       "text11,text12,text13,text14,text15 from candidate where candidate_id= '%s';" \
                                       % (self.Candidate_MS_Id)
            query = self.MS_Cand_Custom_text
            C2.execute_query(query)
            self.ms_data.update({'text1': self.data[0], 'text2': self.data[1],
                                 'text3': self.data[2], 'text4': self.data[3],
                                 'text5': self.data[4],
                                 'text6': self.data[5], 'text7': self.data[6],
                                 'text8': self.data[7],
                                 'text9': self.data[8], 'text10': self.data[9],
                                 'text11': self.data[10],
                                 'text12': self.data[11], 'text13': self.data[12],
                                 'text14': self.data[13],
                                 'text15': self.data[14]})

            self.MS_Cand_Custom_TextArea = "select textarea1,textarea2,textarea3,textarea4 from candidate where " \
                                           "candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Custom_TextArea
            C2.execute_query(query)
            self.ms_data.update({'textArea1': self.data[0], 'textArea2': self.data[1],
                                 'textArea3': self.data[2], 'textArea4': self.data[3]})

            self.MS_Cand_Custom_Integer_values = "select v1.catalog_value_name as drop_down1," \
                                                 "v2.catalog_value_name as drop_down2," \
                                                 "v3.catalog_value_name as drop_down3," \
                                                 "v4.catalog_value_name as drop_down4," \
                                                 "v5.catalog_value_name as drop_down5," \
                                                 "v6.catalog_value_name as drop_down6," \
                                                 "v7.catalog_value_name as drop_down7," \
                                                 "v8.catalog_value_name as drop_down8," \
                                                 "v9.catalog_value_name as drop_down9," \
                                                 "v10.catalog_value_name as drop_down10," \
                                                 "v11.catalog_value_name as drop_down11," \
                                                 "v12.catalog_value_name as drop_down12," \
                                                 "v13.catalog_value_name as drop_down13," \
                                                 "v14.catalog_value_name as drop_down14," \
                                                 "v15.catalog_value_name as drop_down15 from candidate c " \
                                                 "left join catalog_values v1 on c.drop_down1=v1.ams_id " \
                                                 "left join catalog_values v2 on c.drop_down2=v2.ams_id " \
                                                 "left join catalog_values v3 on c.drop_down3=v3.ams_id " \
                                                 "left join catalog_values v4 on c.drop_down4=v4.ams_id " \
                                                 "left join catalog_values v5 on c.drop_down5=v5.ams_id " \
                                                 "left join catalog_values v6 on c.drop_down6=v6.ams_id " \
                                                 "left join catalog_values v7 on c.drop_down7=v7.ams_id " \
                                                 "left join catalog_values v8 on c.drop_down8=v8.ams_id " \
                                                 "left join catalog_values v9 on c.drop_down9=v9.ams_id " \
                                                 "left join catalog_values v10 on c.drop_down10=v10.ams_id " \
                                                 "left join catalog_values v11 on c.drop_down11=v11.ams_id " \
                                                 "left join catalog_values v12 on c.drop_down12=v12.ams_id " \
                                                 "left join catalog_values v13 on c.drop_down13=v13.ams_id " \
                                                 "left join catalog_values v14 on c.drop_down14=v14.ams_id " \
                                                 "left join catalog_values v15 on c.drop_down15=v15.ams_id " \
                                                 "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Custom_Integer_values
            C2.execute_query(query)
            self.ms_data.update({'integer1': self.data[0], 'integer2': self.data[1],
                                 'integer3': self.data[2], 'integer4': self.data[3],
                                 'integer5': self.data[4],
                                 'integer6': self.data[5], 'integer7': self.data[6],
                                 'integer8': self.data[7],
                                 'integer9': self.data[8], 'integer10': self.data[9],
                                 'integer11': self.data[10],
                                 'integer12': self.data[11], 'integer13': self.data[12],
                                 'integer14': self.data[13],
                                 'integer15': self.data[14]})

            self.MS_Cand_Work_Profile = "select c.employer1,c.industry1,c.designation1,c.experience1,c.wp_text1," \
                                        "c.employer2,c.industry2,c.designation2,c.experience2,c.wp_text2," \
                                        "c.employer3,c.industry3,c.designation3,c.experience3,c.wp_text3," \
                                        "c.employer4,c.industry4,c.designation4,c.experience4,c.wp_text4," \
                                        "v1.catalog_value_name as industry1_v,v2.catalog_value_name as industry2_v," \
                                        "v3.catalog_value_name as industry3_v,v4.catalog_value_name as industry4_v " \
                                        "from candidate c left join catalog_values v1 on c.industry1 = v1.ams_id " \
                                        "left join catalog_values v2 on c.industry2 = v2.ams_id " \
                                        "left join catalog_values v3 on c.industry3 = v3.ams_id " \
                                        "left join catalog_values v4 on c.industry4 = v4.ams_id " \
                                        "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Work_Profile
            C2.execute_query(query)

            self.ms_data.update({'employer1': self.data[0], 'industry1': self.data[20],
                                 'designation1': self.data[2], 'experience1': self.data[3],
                                 'wptext1': self.data[4],
                                 'employer2': self.data[5], 'industry2': self.data[21],
                                 'designation2': self.data[7],
                                 'experience2': self.data[8], 'wptext2': self.data[9],
                                 'employer3': self.data[10],
                                 'industry3': self.data[22], 'designation3': self.data[12],
                                 'experience3': self.data[13],
                                 'wptext3': self.data[14], 'employer4': self.data[15],
                                 'industry4': self.data[23], 'designation4': self.data[17],
                                 'experience4': self.data[18],
                                 'wptext4': self.data[19]})

            self.MS_Cand_Tenth_Details = "select c.tenthboard,c.tenthmarks,v.catalog_value_name as tenth_yop " \
                                         "from candidate c left join catalog_values v on c.tenthyop= v.ams_id " \
                                         "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Tenth_Details
            C2.execute_query(query)

            self.ms_data.update({'tenthBoard': self.data[0], 'tenthMark': self.data[1],
                                 'tenthYear': self.data[2]})

            self.MS_Cand_Twelth_Details = " select c.twelthboard,c.twelthmarks,v.catalog_value_name as twelthyop from " \
                                          "candidate c left join catalog_values v on c.twelthyop= v.ams_id " \
                                          "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Twelth_Details
            C2.execute_query(query)

            self.ms_data.update({'twelthBoard': self.data[0], 'twelthMark': self.data[1],
                                 'twelthYear': self.data[2]})

            self.MS_Cand_UG_Details = "select v1.catalog_value_name as ugdegree,v2.catalog_value_name as ugbranch," \
                                      "ugcgpa,v3.catalog_value_name as ugcollege,v4.catalog_value_name as ugyop " \
                                      "from candidate c left join catalog_values v1 on c.ugdegree= v1.ams_id " \
                                      "left join catalog_values v2 on c.ugbranch=  v2.ams_id " \
                                      "left join catalog_values v3 on c.ugcollege= v3.ams_id " \
                                      "left join catalog_values v4 on c.ugyop= v4.ams_id where " \
                                      "candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_UG_Details
            C2.execute_query(query)

            self.ms_data.update({'ugDegree': self.data[0], 'ugBranch': self.data[1],
                                 'ugMark': self.data[2], 'ugCollege': self.data[3],
                                 'ugYop': self.data[4]})

            self.MS_Cand_PG_Details = "select v1.catalog_value_name as pgdegree,v2.catalog_value_name as pgbranch," \
                                      "pgcgpa,v3.catalog_value_name as pgcollege,v4.catalog_value_name as pgyop " \
                                      "from candidate c left join catalog_values v1 on c.pgdegree= v1.ams_id " \
                                      "left join catalog_values v2 on c.pgbranch= v2.ams_id " \
                                      "left join catalog_values v3 on c.pgcollege= v3.ams_id " \
                                      "left join catalog_values v4 on c.pgyop= v4.ams_id " \
                                      "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_PG_Details
            C2.execute_query(query)

            self.ms_data.update({'pgDegree': self.data[0], 'pgBranch': self.data[1],
                                 'pgMark': self.data[2], 'pgCollege': self.data[3],
                                 'pgYop': self.data[4]})

            self.MS_Cand_Others1_Details = "select v1.catalog_value_name as hi1_degree," \
                                           "v2.catalog_value_name as hi1_branch,hi1_cgpa,v3.catalog_value_name as hi1_college" \
                                           ",v4.catalog_value_name as hi1_yop from candidate c " \
                                           "left join catalog_values v1 on c.hi1_degree= v1.ams_id " \
                                           "left join catalog_values v2 on c.hi1_branch= v2.ams_id " \
                                           "left join catalog_values v3 on c.hi1_college= v3.ams_id " \
                                           "left join catalog_values v4 on c.hi1_yop= v4.ams_id " \
                                           "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Others1_Details
            C2.execute_query(query)

            self.ms_data.update({'others1Degree': self.data[0], 'others1Branch': self.data[1],
                                 'others1Mark': self.data[2], 'others1College': self.data[3],
                                 'others1Yop': self.data[4]})

            self.MS_Cand_Others2_Details = "select v1.catalog_value_name as hi2_degree,v2.catalog_value_name as hi2_branch," \
                                           "hi2_cgpa,v3.catalog_value_name as hi2_college, v4.catalog_value_name as hi2_yop " \
                                           "from candidate c left join catalog_values v1 on c.hi2_degree= v1.ams_id " \
                                           "left join catalog_values v2 on c.hi2_branch= v2.ams_id " \
                                           "left join catalog_values v3 on c.hi2_college= v3.ams_id " \
                                           "left join catalog_values v4 on c.hi2_yop= v4.ams_id " \
                                           "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Others2_Details
            C2.execute_query(query)

            self.ms_data.update({'others2Degree': self.data[0], 'others2Branch': self.data[1],
                                 'others2Mark': self.data[2], 'others2College': self.data[3],
                                 'others2Yop': self.data[4]})

            self.MS_Cand_Final_Details = "select v1.catalog_value_name as hi_degree,v2.catalog_value_name as hi_branch," \
                                         "hi_cgpa,v3.catalog_value_name as hi_college,v4.catalog_value_name as hi_yop " \
                                         "from candidate c left join catalog_values v1 on c.hi_degree= v1.ams_id " \
                                         "left join catalog_values v2 on c.hi_branch= v2.ams_id " \
                                         "left join catalog_values v3 on c.hi_college= v3.ams_id " \
                                         "left join catalog_values v4 on c.hi_yop= v4.ams_id " \
                                         "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Final_Details
            C2.execute_query(query)

            self.ms_data.update({'finalDegree': self.data[0], 'finalBranch': self.data[1],
                                 'finalMark': self.data[2], 'finalCollege': self.data[3],
                                 'finalYop': self.data[4]})

            self.MS_Cand_Degree_Infos = "SELECT cv1.catalog_value_name AS tenthboard2," \
                                        "cv2.catalog_value_name AS tenthcity," \
                                        "cv3.catalog_value_name AS tenthstate, tenthmarksoutof, " \
                                        "cv4.catalog_value_name AS twelthboard2, cv5.catalog_value_name AS twelthstate, " \
                                        "cv6.catalog_value_name AS twelthcity, twelthmarksoutof, " \
                                        "cv7.catalog_value_name AS uguniversity, cv8.catalog_value_name AS pguniversity, " \
                                        "cv9.catalog_value_name AS hiuniversity, cv10.catalog_value_name AS hi2university, " \
                                        "cv11.catalog_value_name AS hi1university, cv12.catalog_value_name AS ugcity, " \
                                        "cv13.catalog_value_name AS pgcity, cv14.catalog_value_name AS hicity, " \
                                        "cv15.catalog_value_name AS hi2city, cv16.catalog_value_name AS hi1city, " \
                                        "cv17.catalog_value_name AS ugstate, cv18.catalog_value_name AS pgstate, " \
                                        "cv19.catalog_value_name AS histate, cv20.catalog_value_name AS hi2state, " \
                                        "cv21.catalog_value_name AS hi1state, ugmarksoutof, pgmarksoutof, " \
                                        "himarksoutof, hi2marksoutof, hi1marksoutof FROM candidate c " \
                                        "LEFT JOIN catalog_values cv1 ON c.tenthboard2 = cv1.ams_id " \
                                        "LEFT JOIN catalog_values cv2 ON c.tenthcity = cv2.ams_id " \
                                        "LEFT JOIN catalog_values cv3 ON c.tenthstate = cv3.ams_id " \
                                        "LEFT JOIN catalog_values cv4 ON c.twelthboard2 = cv4.ams_id " \
                                        "LEFT JOIN catalog_values cv5 ON c.twelthstate = cv5.ams_id " \
                                        "LEFT JOIN catalog_values cv6 ON c.twelthcity = cv6.ams_id " \
                                        "LEFT JOIN catalog_values cv7 ON c.uguniversity = cv7.ams_id " \
                                        "LEFT JOIN catalog_values cv8 ON c.pguniversity = cv8.ams_id " \
                                        "LEFT JOIN catalog_values cv9 ON c.hiuniversity = cv9.ams_id " \
                                        "LEFT JOIN catalog_values cv10 ON c.hi2university = cv10.ams_id " \
                                        "LEFT JOIN catalog_values cv11 ON c.hi1university = cv11.ams_id " \
                                        "LEFT JOIN catalog_values cv12 ON c.ugcity = cv12.ams_id " \
                                        "LEFT JOIN catalog_values cv13 ON c.pgcity = cv13.ams_id " \
                                        "LEFT JOIN catalog_values cv14 ON c.hicity = cv14.ams_id " \
                                        "LEFT JOIN catalog_values cv15 ON c.hi2city = cv15.ams_id " \
                                        "LEFT JOIN catalog_values cv16 ON c.hi1city = cv16.ams_id " \
                                        "LEFT JOIN catalog_values cv17 ON c.ugstate = cv17.ams_id " \
                                        "LEFT JOIN catalog_values cv18 ON c.pgstate = cv18.ams_id " \
                                        "LEFT JOIN catalog_values cv19 ON c.histate = cv19.ams_id " \
                                        "LEFT JOIN catalog_values cv20 ON c.hi2state = cv20.ams_id " \
                                        "LEFT JOIN catalog_values cv21 ON c.hi1state = cv21.ams_id " \
                                        "where candidate_id= '%s';" % (self.Candidate_MS_Id)
            query = self.MS_Cand_Degree_Infos
            C2.execute_query(query)

            self.ms_data.update({'tenthBoard2': self.data[0], 'tenthCity': self.data[1],
                                 'tenthState': self.data[2], 'tenthOutOf': self.data[3],
                                 'twelthBoard2': self.data[4],
                                 'twelthState': self.data[5], 'twelthCity': self.data[6],
                                 'twelthOutOf': self.data[7],
                                 'ugUniversity': self.data[8], 'pgUniversity': self.data[9],
                                 'finalUniversity': self.data[10],
                                 'others2University': self.data[11],
                                 'others1University': self.data[12],
                                 'ugCity': self.data[13], 'pgCity': self.data[14],
                                 'finalCity': self.data[15], 'others2City': self.data[16],
                                 'others1City': self.data[17], 'ugState': self.data[18],
                                 'pgState': self.data[19], 'finalState': self.data[20],
                                 'others2State': self.data[21], 'others1State': self.data[22],
                                 'ugOutOf': self.data[23], 'pgOutOf': self.data[24],
                                 'finalOutOf': self.data[25], 'others2OutOf': self.data[26],
                                 'others1OutOf': self.data[27]})

        else:
            print ("Candidate is not created in microsite")
            self.Candidate_MS_Id = None
            self.candidate_amsid = None

        self.conn1.close()

    def ams_query(self):

        C2.amsdbconnection()
        self.ams_data = {}
        self.candidate_ams_id = self.candidate_amsid
        self.ams_cand_personal = "select c.first_name,c.middle_name,c.last_name,c.email1,c.email2,c.usn,c.Mobile1," \
                                 "c.phone_office,c.passport,c.pan_card,c.marital_status,cv.value as nationality," \
                                 "gender,aadhaar_no,current_ctc " \
                                 "from candidates c left join catalog_values cv on c.nationality = cv.id " \
                                 "where c.id= '%s';" % (self.candidate_ams_id)
        query = self.ams_cand_personal
        C2.execute_query(query)

        self.ams_data.update(
            {'firstName': self.data[0], 'middleName': self.data[1], 'lastName': self.data[2], 'email': self.data[3],
             'altEmail': self.data[4], 'usn': self.data[5], 'mobile': self.data[6],
             'altMobile': self.data[7], 'passport': self.data[8], 'pancard': self.data[9],
             'maritalStatus': self.data[10], 'nationality': self.data[11], 'gender': self.data[12],
             'aadhar': self.data[13], 'currentCtc': self.data[14]})

        self.ams_cand_work_experience = "select current_experience,bpo_experience,total_experience,current_employer_text " \
                                        "from candidates where id= '%s';" % (self.candidate_ams_id)
        query = self.ams_cand_work_experience
        C2.execute_query(query)

        self.ams_data.update(
            {'currentExperience': self.data[0], 'bpoExperience': self.data[1], 'totalExperience': self.data[2],
             'currentCompany': self.data[3]})

        self.ams_cand_address = "select address1,address2 from candidates where id= '%s';" % (self.candidate_ams_id)
        query = self.ams_cand_address
        C2.execute_query(query)
        res = list(self.data)
        self.current_address = [str(s).strip() for s in res[0].split(',')]
        self.permanent_address = [str(s).strip() for s in res[1].split(',')]

        self.ams_data.update(
            {'currentAddress1': self.current_address[0], 'currentAddress2': self.current_address[1],
             'currentCity': self.current_address[2],
             'currentState': self.current_address[3],
             'currentPincode': self.current_address[4],
             'permanentAddress1': self.permanent_address[0], 'permanentAddress2': self.permanent_address[1],
             'permanentCity': self.permanent_address[2], 'permanentState': self.permanent_address[3],
             'country': self.permanent_address[4],
             'permanentPincode': self.permanent_address[5]})

        self.ams_cand_custom_text = "select c.text1,c.Text2,c.Text3,c.Text4,c.Text5,cs.text6,cs.Text7," \
                                    "cs.Text8,cs.Text9,cs.Text10,text11,text12,text13,text14,text15 from candidates c " \
                                    "left join candidate_customs cs on c.candidatecustom_id = cs.id " \
                                    "where c.id= '%s';" % (self.candidate_ams_id)
        query = self.ams_cand_custom_text
        C2.execute_query(query)

        self.ams_data.update({'text1': self.data[0], 'text2': self.data[1],
                              'text3': self.data[2], 'text4': self.data[3],
                              'text5': self.data[4],
                              'text6': self.data[5], 'text7': self.data[6],
                              'text8': self.data[7],
                              'text9': self.data[8], 'text10': self.data[9],
                              'text11': self.data[10],
                              'text12': self.data[11], 'text13': self.data[12],
                              'text14': self.data[13],
                              'text15': self.data[14]})

        self.ams_cand_custom_textarea = "select text_area1,text_area2,text_area3,text_area4 " \
                                        "from candidates where id= '%s';" % (self.candidate_ams_id)
        query = self.ams_cand_custom_textarea
        C2.execute_query(query)

        self.ams_data.update({'textArea1': self.data[0], 'textArea2': self.data[1],
                              'textArea3': self.data[2], 'textArea4': self.data[3]})

        self.ams_cand_custom_integer = "select v1.value as integer1,v2.value as integer2,v3.value as integer3," \
                                       "v4.value integer4,v5.value as integer5,v6.value as integer6,v7.value as integer7," \
                                       "v8.value as integer8,v9.value as integer9,v10.value as integer10," \
                                       "v11.value as integer11,v12.value as integer12,v13.value as integer13," \
                                       "v14.value as integer14,v15.value as integer15 " \
                                       "from candidates c left join candidate_customs cs on c.candidatecustom_id = cs.id " \
                                       "left join catalog_values v1 on c.integer1=v1.id " \
                                       "left join catalog_values v2 on c.integer2=v2.id " \
                                       "left join catalog_values v3 on c.integer3=v3.id " \
                                       "left join catalog_values v4 on c.integer4=v4.id " \
                                       "left join catalog_values v5 on c.integer5=v5.id " \
                                       "left join catalog_values v6 on cs.integer6=v6.id " \
                                       "left join catalog_values v7 on cs.integer7=v7.id " \
                                       "left join catalog_values v8 on cs.integer8=v8.id " \
                                       "left join catalog_values v9 on cs.integer9=v9.id " \
                                       "left join catalog_values v10 on cs.integer10=v10.id " \
                                       "left join catalog_values v11 on cs.integer11=v11.id " \
                                       "left join catalog_values v12 on cs.integer12=v12.id " \
                                       "left join catalog_values v13 on cs.integer13=v13.id " \
                                       "left join catalog_values v14 on cs.integer14=v14.id " \
                                       "left join catalog_values v15 on cs.integer15=v15.id " \
                                       "where c.id= '%s';" % (self.candidate_ams_id)
        query = self.ams_cand_custom_integer
        C2.execute_query(query)

        self.ams_data.update({'integer1': self.data[0], 'integer2': self.data[1],
                              'integer3': self.data[2], 'integer4': self.data[3],
                              'integer5': self.data[4],
                              'integer6': self.data[5], 'integer7': self.data[6],
                              'integer8': self.data[7],
                              'integer9': self.data[8], 'integer10': self.data[9],
                              'integer11': self.data[10],
                              'integer12': self.data[11], 'integer13': self.data[12],
                              'integer14': self.data[13],
                              'integer15': self.data[14]})

        self.ams_cand_education = "select ifnull(c.degree_text,v2.value) as degree," \
                                  "ifnull(c.degree_type_text,v3.value) as branch," \
                                  "ifnull(c.college_text,v1.value) as college,v4.value as yop,c.percentage," \
                                  "ifnull(c.university_text,v5.value)as university,v6.value as city," \
                                  "v7.value as state,c.percentage_out_of,v8.value as board" \
                                  " from candidate_education_profiles c left join catalog_values v1 on c.college_id=v1.id " \
                                  "left join catalog_values v2 on c.degree_id=v2.id left join catalog_values v3 on c.degree_type_id=v3.id " \
                                  "left join catalog_values v4 on c.end_year=v4.id " \
                                  "left join catalog_values v5 on c.university_id=v5.id " \
                                  "left join catalog_values v6 on c.city_id=v6.id " \
                                  "left join catalog_values v7 on c.state_id=v7.id " \
                                  "left join catalog_values v8 on c.board_id=v8.id " \
                                  "where candidate_id = '%s' order by yop asc;" % (self.candidate_ams_id)
        query = self.ams_cand_education
        # print query
        C2.execute_query2(query)
        self.ams_cand_tenth = [str(s).strip() for s in list(self.data[0])]
        self.ams_cand_twelth = [str(s).strip() for s in list(self.data[1])]
        self.ams_cand_ug = [str(s).strip() for s in list(self.data[2])]
        self.ams_cand_pg = [str(s).strip() for s in list(self.data[3])]
        self.ams_cand_others1 = [str(s).strip() for s in list(self.data[4])]
        self.ams_cand_others2 = [str(s).strip() for s in list(self.data[5])]
        self.ams_cand_final = [str(s).strip() for s in list(self.data[6])]

        self.ams_data.update(
            {'tenthBoard': self.ams_cand_tenth[2], 'tenthYear': self.ams_cand_tenth[3],
             'tenthMark': int(float(self.ams_cand_tenth[4])), 'tenthCity': self.ams_cand_tenth[6],
             'tenthState': self.ams_cand_tenth[7], 'tenthOutOf': int(float(self.ams_cand_tenth[8])),
             'tenthBoard2': self.ams_cand_tenth[9]})

        self.ams_data.update(
            {'twelthBoard': self.ams_cand_twelth[2], 'twelthYear': self.ams_cand_twelth[3],
             'twelthMark': int(float(self.ams_cand_twelth[4])), 'twelthCity': self.ams_cand_twelth[6],
             'twelthState': self.ams_cand_twelth[7], 'twelthOutOf': int(float(self.ams_cand_twelth[8])),
             'twelthBoard2': self.ams_cand_twelth[9]})

        self.ams_data.update(
            {'ugDegree': self.ams_cand_ug[0], 'ugBranch': self.ams_cand_ug[1],
             'ugCollege': self.ams_cand_ug[2], 'ugYop': self.ams_cand_ug[3],
             'ugMark': int(float(self.ams_cand_ug[4])), 'ugCity': self.ams_cand_ug[6],
             'ugState': self.ams_cand_ug[7], 'ugOutOf': int(float(self.ams_cand_ug[8])),
             'ugUniversity': self.ams_cand_ug[5]})

        self.ams_data.update(
            {'pgDegree': self.ams_cand_pg[0], 'pgBranch': self.ams_cand_pg[1],
             'pgCollege': self.ams_cand_pg[2], 'pgYop': self.ams_cand_pg[3],
             'pgMark': int(float(self.ams_cand_pg[4])), 'pgCity': self.ams_cand_pg[6],
             'pgState': self.ams_cand_pg[7], 'pgOutOf': int(float(self.ams_cand_pg[8])),
             'pgUniversity': self.ams_cand_pg[5]})

        self.ams_data.update(
            {'others1Degree': self.ams_cand_others1[0], 'others1Branch': self.ams_cand_others1[1],
             'others1College': self.ams_cand_others1[2], 'others1Yop': self.ams_cand_others1[3],
             'others1Mark': int(float(self.ams_cand_others1[4])), 'others1City': self.ams_cand_others1[6],
             'others1State': self.ams_cand_others1[7], 'others1OutOf': int(float(self.ams_cand_others1[8])),
             'others1University': self.ams_cand_others1[5]})

        self.ams_data.update(
            {'others2Degree': self.ams_cand_others2[0], 'others2Branch': self.ams_cand_others2[1],
             'others2College': self.ams_cand_others2[2], 'others2Yop': self.ams_cand_others2[3],
             'others2Mark': int(float(self.ams_cand_others2[4])), 'others2City': self.ams_cand_others2[6],
             'others2State': self.ams_cand_others2[7], 'others2OutOf': int(float(self.ams_cand_others2[8])),
             'others2University': self.ams_cand_others2[5]})


        self.ams_data.update(
            {'finalDegree': self.ams_cand_final[0], 'finalBranch': self.ams_cand_final[1],
             'finalCollege': self.ams_cand_final[2], 'finalYop': self.ams_cand_final[3],
             'finalMark': int(float(self.ams_cand_final[4])), 'finalCity': self.ams_cand_final[6],
             'finalState': self.ams_cand_final[7], 'finalOutOf': int(float(self.ams_cand_final[8])),
             'finalUniversity': self.ams_cand_final[5]})

        self.ams_cand_work_profile = "select designation_text,employer_text,cv.value as industry,title,experience from candidate_work_profiles c " \
                                     "left join catalog_values cv on c.industry =cv.id where candidate_id= '%s';" % (
                                         self.candidate_ams_id)
        query = self.ams_cand_work_profile
        C2.execute_query2(query)
        self.ams_cand_wp4 = [str(s).strip() for s in list(self.data[0])]
        self.ams_cand_wp3 = [str(s).strip() for s in list(self.data[1])]
        self.ams_cand_wp2 = [str(s).strip() for s in list(self.data[2])]
        self.ams_cand_wp1 = [str(s).strip() for s in list(self.data[3])]

        self.ams_data.update(
            {'designation1': self.ams_cand_wp1[0], 'employer1': self.ams_cand_wp1[1],
             'industry1': self.ams_cand_wp1[2], 'wptext1': self.ams_cand_wp1[3],
             'experience1': self.ams_cand_wp1[4]})

        self.ams_data.update(
            {'designation2': self.ams_cand_wp2[0], 'employer2': self.ams_cand_wp2[1],
             'industry2': self.ams_cand_wp2[2], 'wptext2': self.ams_cand_wp2[3],
             'experience2': self.ams_cand_wp2[4]})

        self.ams_data.update(
            {'designation3': self.ams_cand_wp3[0], 'employer3': self.ams_cand_wp3[1],
             'industry3': self.ams_cand_wp3[2], 'wptext3': self.ams_cand_wp3[3],
             'experience3': self.ams_cand_wp3[4]})

        self.ams_data.update(
            {'designation4': self.ams_cand_wp4[0], 'employer4': self.ams_cand_wp4[1],
             'industry4': self.ams_cand_wp4[2], 'wptext4': self.ams_cand_wp4[3],
             'experience4': self.ams_cand_wp4[4]})

        self.conn.close()

    def execute_query(self, query):
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchone()
        except Exception as e:
            print (e)
            print("Execute_Query Method is Throwing Exception")

    def execute_query2(self, query):
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchall()
        except:
            print("Execute_Query2 Method is Throwing Exception")

    def compare_values1(self, ms_data, ui_data, ams_data, mess):
        if (ms_data == ui_data) and (ms_data == ams_data):
            self.ws.write(self.rowsize, self.a, ui_data, self.style1)
            self.ws.write(self.rowsize + 1, self.a, ms_data, self.style3)
            self.ws.write(self.rowsize + 2, self.a, ams_data, self.style3)
            self.a = self.a + 1
        elif (ms_data == ui_data) and (ms_data != ams_data):
            self.ms_ams_row_status = 'Fail'
            self.ms_ams_status_color = self.style2
            self.ws.write(self.rowsize, self.a, ui_data, self.style1)
            self.ws.write(self.rowsize + 1, self.a, ms_data, self.style3)
            self.ws.write(self.rowsize + 2, self.a, ams_data, self.style2)
            self.a = self.a + 1
        else:
            self.ui_ms_row_status = 'Fail'
            self.ui_ms_status_color = self.style2
            self.ws.write(self.rowsize, self.a, ui_data, self.style1)
            self.ws.write(self.rowsize + 1, self.a, ms_data, self.style2)
            self.ws.write(self.rowsize + 2, self.a, ams_data, self.style3)
            self.a = self.a + 1

    def write_header(self, value):
        totallen = len(value)
        print (totallen)
        # self.ws.write(0, 0, "Overall Status", self.__style3)
        for i in range(0, totallen):
            self.ws.write(self.rowsize, i, value[i], self.style0)

    def output_excel_header(self):
        self.rowsize = 1
        self.col1 = 0

        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Extract Resume Result')

        a = ['Data', 'Status', 'First Name', 'Middle Name', 'Last Name', 'Email', 'Alternate Email', 'Mobile',
             'Alternate mobile', 'Pancard', 'Passport', 'Gender', 'Marital', 'Nationality',

             'Tenth Board', 'Tenth YOP', 'Tenth Mark', 'Twelth Board', 'Twelth YOP', 'Twelth Mark',
             'UG College', 'UG Branch', 'UG Degree', 'UG YOP', 'UG Percentage',
             'PG College', 'PG Branch', 'PG Degree', 'PG YOP', 'PG Percentage',
             'Others1 College', 'Others1 Branch', 'Others1 Degree', 'Others1 YOP', 'Others1 Percentage',
             'Others2 College', 'Others2 Branch', 'Others2 Degree', 'Others2 YOP', 'Others2 Percentage',
             'Final College', 'Final Branch', 'Final Degree', 'Final YOP', 'Final Percentage',

             'Current Address1', 'Current Address2', 'Current City', 'Current State', 'Current Pincode',
             'Permanent Address1', 'Permanent Address2', 'Permanent City', 'Permanent State', 'Permanent Pincode',
             'Current Company', 'Current Experience', 'Total Experience', 'Bpo Experience',

             'Text1', 'Text2', 'Text3', 'Text4', 'Text5', 'Text6', 'Text7', 'Text8', 'Text9', 'Text10',
             'Textarea1', 'Textarea2', 'Textarea3', 'Textarea4', 'Integer1', 'Integer2', 'Integer3', 'Integer4',
             'Integer5', 'Integer6', 'Integer7', 'Integer8', 'Integer9', 'Integer10',

             'Employee1', 'Industry1', 'Experience1', 'Designation1', 'wptext1',
             'Employee2', 'Industry2', 'Experience2', 'Designation2', 'wptext2',
             'Employee3', 'Industry3', 'Experience3', 'Designation3', 'wptext3',
             'Employee4', 'Industry4', 'Experience4', 'Designation4', 'wptext4', 'USN', 'Aadhar', 'Current CTC',
             'Tenth Board2', 'Tenth City', 'Tenth State', 'Tenth Out of', 'Twelth Board2',
             'TwelthCity', 'Twelth State', 'Twleth Out of',
             'UG University', 'UG City', 'UG State', 'UG Out of',
             'PG University', 'PG City', 'PG State', 'PG Out of',
             'Hi1 University', 'Hi1 City', 'Hi1 State', 'Hi1 Out of', 'Hi2 University', 'Hi2 City',
             'Hi2 State', 'Hi2 Out of', 'Final University', 'Final City', 'Final State', 'Final Out of',
             'Total Experience in Years', 'Text11', 'Text12', 'Text13',
             'Text14', 'Text15', 'Integer11', 'Integer12', 'Integer13', 'Integer14', 'Integer15', 'Country']
        self.write_header(a)

    def validation_ms_and_ui(self):
        self.a = 2
        self.rowsize += 2
        self.ui_ms_row_status = 'Pass'
        self.ui_ms_status_color = self.style3
        self.ms_ams_row_status = 'Pass'
        self.ms_ams_status_color = self.style3
        # print "Compare method is started"
        self.ws.write(self.rowsize, 0, "UI Input", self.style0)
        self.ws.write(self.rowsize + 1, 0, "MS DB Data", self.style0)
        self.ws.write(self.rowsize + 2, 0, "AMS DB Data", self.style0)
        self.compare_values1(self.ms_data.get('firstName'), C2.ui_fname, self.ams_data.get('firstName'),
                             'First name is  not matched')
        self.compare_values1(self.ms_data.get('middleName'), C2.ui_mname, self.ams_data.get('middleName'),
                             'Middle name is  not matched')
        self.compare_values1(self.ms_data.get('lastName'), C2.ui_lname, self.ams_data.get('lastName'),
                             'Last name is not matched')
        self.compare_values1(self.ms_data.get('email'), C2.ui_email, self.ams_data.get('email'), 'Email is not matched')
        self.compare_values1(self.ms_data.get('altEmail'), C2.ui_altmail, self.ams_data.get('altEmail'),
                             'Alternate Email is not matched')
        self.compare_values1(self.ms_data.get('mobile'), C2.ui_mobile, self.ams_data.get('mobile'),
                             'Mobile is not matched')
        self.compare_values1(self.ms_data.get('altMobile'), C2.ui_altmobile, self.ams_data.get('altMobile'),
                             'Alternate Mobile is not matched')
        self.compare_values1(self.ms_data.get('pancard'), C2.ui_pancard, self.ams_data.get('pancard'),
                             'Pancard is not matched')
        self.compare_values1(self.ms_data.get('passport'), C2.ui_passport, self.ams_data.get('passport'),
                             'Passport is not matched')
        self.compare_values1(self.ms_data.get('gender'), C2.ui_gender, self.ams_data.get('gender'),
                             'Gender is not matched')
        self.compare_values1(self.ms_data.get('maritalStatus'), C2.ui_marital, self.ams_data.get('maritalStatus'),
                             'Marital is not matched')
        self.compare_values1(self.ms_data.get('nationality'), C2.ui_nationality, self.ams_data.get('nationality'),
                             'Nationality is not matched')

        self.compare_values1(self.ms_data.get('tenthBoard'), C2.ui_tenthboard, self.ams_data.get('tenthBoard'),
                             'Tenth Board is not matched')
        self.compare_values1(self.ms_data.get('tenthYear'), C2.ui_tenthyop, self.ams_data.get('tenthYear'),
                             'Tenth Year is not matched')
        self.compare_values1(self.ms_data.get('tenthMark'), C2.ui_tenthmark, str(self.ams_data.get('tenthMark')),
                             'Tenth Mark is not matched')

        self.compare_values1(self.ms_data.get('twelthBoard'), C2.ui_twelthboard, self.ams_data.get('twelthBoard'),
                             'Twelth Board is not matched')
        self.compare_values1(self.ms_data.get('twelthYear'), C2.ui_twelthyop, self.ams_data.get('twelthYear'),
                             'Twelth Year is not matched')
        self.compare_values1(self.ms_data.get('twelthMark'), C2.ui_twelthmark, str(self.ams_data.get('twelthMark')),
                             'Twelth Mark is not matched')

        self.compare_values1(self.ms_data.get('ugCollege'), C2.ui_ugcollege, self.ams_data.get('ugCollege'),
                             'UG College is not matched')
        self.compare_values1(self.ms_data.get('ugBranch'), C2.ui_ugbranch, self.ams_data.get('ugBranch'),
                             'UG Branch is not matched')
        self.compare_values1(self.ms_data.get('ugDegree'), C2.ui_ugdegree, self.ams_data.get('ugDegree'),
                             'UG Degree is not matched')
        self.compare_values1(self.ms_data.get('ugYop'), C2.ui_ugyop, self.ams_data.get('ugYop'),
                             'UG YOP is not matched')
        self.compare_values1(self.ms_data.get('ugMark'), C2.ui_ugmark, str(self.ams_data.get('ugMark')),
                             'UG Percentage is not matched')

        self.compare_values1(self.ms_data.get('pgCollege'), C2.ui_pgcollege, self.ams_data.get('pgCollege'),
                             'PG College is not matched')
        self.compare_values1(self.ms_data.get('pgBranch'), C2.ui_pgbranch, self.ams_data.get('pgBranch'),
                             'PG Branch is not matched')
        self.compare_values1(self.ms_data.get('pgDegree'), C2.ui_pgdegree, self.ams_data.get('pgDegree'),
                             'PG Degree is not matched')
        self.compare_values1(self.ms_data.get('pgYop'), C2.ui_pgyop, self.ams_data.get('pgYop'),
                             'PG YOP is not matched')
        self.compare_values1(self.ms_data.get('pgMark'), C2.ui_pgmark, str(self.ams_data.get('pgMark')),
                             'PG Percentage is not matched')

        self.compare_values1(self.ms_data.get('others1College'), C2.ui_others1college,
                             self.ams_data.get('others1College'), 'Others1 College is not matched')
        self.compare_values1(self.ms_data.get('others1Branch'), C2.ui_others1branch, self.ams_data.get('others1Branch'),
                             'Others1 Branch is not matched')
        self.compare_values1(self.ms_data.get('others1Degree'), C2.ui_others1degree, self.ams_data.get('others1Degree'),
                             'Others1 Degree is not matched')
        self.compare_values1(self.ms_data.get('others1Yop'), C2.ui_others1yop, self.ams_data.get('others1Yop'),
                             'Others1 YOP is not matched')
        self.compare_values1(self.ms_data.get('others1Mark'), C2.ui_others1mark, str(self.ams_data.get('others1Mark')),
                             'Others1 Percentage is not matched')

        self.compare_values1(self.ms_data.get('others2College'), C2.ui_others2college,
                             self.ams_data.get('others2College'), 'Others2 College is not matched')
        self.compare_values1(self.ms_data.get('others2Branch'), C2.ui_others2branch, self.ams_data.get('others2Branch'),
                             'Others2 Branch is not matched')
        self.compare_values1(self.ms_data.get('others2Degree'), C2.ui_others2degree, self.ams_data.get('others2Degree'),
                             'Others2 Degree is not matched')
        self.compare_values1(self.ms_data.get('others2Yop'), C2.ui_others2yop, self.ams_data.get('others2Yop'),
                             'Others2 YOP is not matched')
        self.compare_values1(self.ms_data.get('others2Mark'), C2.ui_others2mark, str(self.ams_data.get('others2Mark')),
                             'Others2 Percentage is not matched')

        self.compare_values1(self.ms_data.get('finalCollege'), C2.ui_finalcollege, self.ams_data.get('finalCollege'),
                             'Final College is not matched')
        self.compare_values1(self.ms_data.get('finalBranch'), C2.ui_finalbranch, self.ams_data.get('finalBranch'),
                             'Final Branch is not matched')
        self.compare_values1(self.ms_data.get('finalDegree'), C2.ui_finaldegree, self.ams_data.get('finalDegree'),
                             'final Degree is not matched')
        self.compare_values1(self.ms_data.get('finalYop'), C2.ui_finalyop, self.ams_data.get('finalYop'),
                             'final YOP is not matched')
        self.compare_values1(self.ms_data.get('finalMark'), C2.ui_finalmark, str(self.ams_data.get('finalMark')),
                             'Final Percentage is not matched')

        self.compare_values1(self.ms_data.get('currentAddress1'), C2.ui_ca1, self.ams_data.get('currentAddress1'),
                             'CurrentAddress1 is not matched')
        self.compare_values1(self.ms_data.get('currentAddress2'), C2.ui_ca2, self.ams_data.get('currentAddress2'),
                             'CurrentAddress2 is not matched')
        self.compare_values1(self.ms_data.get('currentCity'), C2.ui_ccity, self.ams_data.get('currentCity'),
                             'Current City is not matched')
        self.compare_values1(self.ms_data.get('currentState'), C2.ui_cstate, self.ams_data.get('currentState'),
                             'Current State is not matched')
        self.compare_values1(self.ms_data.get('currentPincode'), C2.ui_cpincode, self.ams_data.get('currentPincode'),
                             'Current Pincode is not matched')

        self.compare_values1(self.ms_data.get('permanentAddress1'), C2.ui_pa1, self.ams_data.get('permanentAddress1'),
                             'Permanent Address1 is not matched')
        self.compare_values1(self.ms_data.get('permanentAddress2'), C2.ui_pa2, self.ams_data.get('permanentAddress2'),
                             'Permanent Address2 is not matched')
        self.compare_values1(self.ms_data.get('permanentCity'), C2.ui_pcity, self.ams_data.get('permanentCity'),
                             'permanent City is not matched')
        self.compare_values1(self.ms_data.get('permanentState'), C2.ui_pstate, self.ams_data.get('permanentState'),
                             'permanent State is not matched')
        self.compare_values1(self.ms_data.get('permanentPincode'), C2.ui_ppincode,
                             self.ams_data.get('permanentPincode'), 'permanent Pincode is not matched')

        self.compare_values1(self.ms_data.get('currentCompany'), C2.ui_ccompany, self.ams_data.get('currentCompany'),
                             'Current company is not matched')
        self.compare_values1(self.ms_data.get('currentExperience'), C2.ui_cexp,
                             str(self.ams_data.get('currentExperience')),
                             'Current Experience is not matched')
        self.compare_values1(self.ms_data.get('totalExperience'), C2.ui_texp, str(self.ams_data.get('totalExperience')),
                             'Total Experience is not matched')
        self.compare_values1(self.ms_data.get('bpoExperience'), C2.ui_bpoexp, str(self.ams_data.get('bpoExperience')),
                             'BPO Experience is not matched')

        self.compare_values1(self.ms_data.get('text1'), C2.ui_text1, self.ams_data.get('text1'), 'Text1 is not matched')
        self.compare_values1(self.ms_data.get('text2'), C2.ui_text2, self.ams_data.get('text2'), 'Text2 is not matched')
        self.compare_values1(self.ms_data.get('text3'), C2.ui_text3, self.ams_data.get('text3'), 'Text3 is not matched')
        self.compare_values1(self.ms_data.get('text4'), C2.ui_text4, self.ams_data.get('text4'), 'Text4 is not matched')
        self.compare_values1(self.ms_data.get('text5'), C2.ui_text5, self.ams_data.get('text5'), 'Text5 is not matched')
        self.compare_values1(self.ms_data.get('text6'), C2.ui_text6, self.ams_data.get('text6'), 'Text6 is not matched')
        self.compare_values1(self.ms_data.get('text7'), C2.ui_text7, self.ams_data.get('text7'), 'Text7 is not matched')
        self.compare_values1(self.ms_data.get('text8'), C2.ui_text8, self.ams_data.get('text8'), 'Text8 is not matched')
        self.compare_values1(self.ms_data.get('text9'), C2.ui_text9, self.ams_data.get('text9'), 'Text9 is not matched')
        self.compare_values1(self.ms_data.get('text10'), C2.ui_text10, self.ams_data.get('text10'),
                             'Text10 is not matched')

        self.compare_values1(self.ms_data.get('textArea1'), C2.ui_textarea1, self.ams_data.get('textArea1'),
                             'TextArea1 is not matched')
        self.compare_values1(self.ms_data.get('textArea2'), C2.ui_textarea2, self.ams_data.get('textArea2'),
                             'TextArea2 is not matched')
        self.compare_values1(self.ms_data.get('textArea3'), C2.ui_textarea3, self.ams_data.get('textArea3'),
                             'TextArea3 is not matched')
        self.compare_values1(self.ms_data.get('textArea4'), C2.ui_textarea4, self.ams_data.get('textArea4'),
                             'TextArea4 is not matched')

        self.compare_values1(self.ms_data.get('integer1'), C2.ui_integer1, self.ams_data.get('integer1'),
                             'Integer1 is not matched')
        self.compare_values1(self.ms_data.get('integer2'), C2.ui_integer2, self.ams_data.get('integer2'),
                             'Integer2 is not matched')
        self.compare_values1(self.ms_data.get('integer3'), C2.ui_integer3, self.ams_data.get('integer3'),
                             'Integer3 is not matched')
        self.compare_values1(self.ms_data.get('integer4'), C2.ui_integer4, self.ams_data.get('integer4'),
                             'Integer4 is not matched')
        self.compare_values1(self.ms_data.get('integer5'), C2.ui_integer5, self.ams_data.get('integer5'),
                             'Integer5 is not matched')
        self.compare_values1(self.ms_data.get('integer6'), C2.ui_integer6, self.ams_data.get('integer6'),
                             'Integer6 is not matched')
        self.compare_values1(self.ms_data.get('integer7'), C2.ui_integer7, self.ams_data.get('integer7'),
                             'Integer7 is not matched')
        self.compare_values1(self.ms_data.get('integer8'), C2.ui_integer8, self.ams_data.get('integer8'),
                             'Integer8 is not matched')
        self.compare_values1(self.ms_data.get('integer9'), C2.ui_integer9, self.ams_data.get('integer9'),
                             'Integer9 is not matched')
        self.compare_values1(self.ms_data.get('integer10'), C2.ui_integer10, self.ams_data.get('integer10'),
                             'Integer10 is not matched')

        self.compare_values1(self.ms_data.get('employer1'), C2.ui_emp1, self.ams_data.get('employer1'),
                             'Employer1 is not matched')
        self.compare_values1(self.ms_data.get('industry1'), C2.ui_indus1, self.ams_data.get('industry1'),
                             'Industry1 is not matched')
        self.compare_values1(self.ms_data.get('experience1'), C2.ui_expe1, int(self.ams_data.get('experience1')),
                             'Experience1 is not matched')
        self.compare_values1(self.ms_data.get('designation1'), C2.ui_desg1, self.ams_data.get('designation1'),
                             'Designation1 is not matched')
        self.compare_values1(self.ms_data.get('wptext1'), C2.ui_wpt1, self.ams_data.get('wptext1'),
                             'wptext1 is not matched')

        self.compare_values1(self.ms_data.get('employer2'), C2.ui_emp2, self.ams_data.get('employer2'),
                             'Employer2 is not matched')
        self.compare_values1(self.ms_data.get('industry2'), C2.ui_indus2, self.ams_data.get('industry2'),
                             'Industry2 is not matched')
        self.compare_values1(self.ms_data.get('experience2'), C2.ui_expe2, int(self.ams_data.get('experience2')),
                             'Experience2 is not matched')
        self.compare_values1(self.ms_data.get('designation2'), C2.ui_desg2, self.ams_data.get('designation2'),
                             'Designation2 is not matched')
        self.compare_values1(self.ms_data.get('wptext2'), C2.ui_wpt2, self.ams_data.get('wptext2'),
                             'wptext2 is not matched')

        self.compare_values1(self.ms_data.get('employer3'), C2.ui_emp3, self.ams_data.get('employer3'),
                             'Employer3 is not matched')
        self.compare_values1(self.ms_data.get('industry3'), C2.ui_indus3, self.ams_data.get('industry3'),
                             'Industry3 is not matched')
        self.compare_values1(self.ms_data.get('experience3'), C2.ui_expe3, int(self.ams_data.get('experience3')),
                             'Experience3 is not matched')
        self.compare_values1(self.ms_data.get('designation3'), C2.ui_desg3, self.ams_data.get('designation3'),
                             'Designation3 is not matched')
        self.compare_values1(self.ms_data.get('wptext3'), C2.ui_wpt3, self.ams_data.get('wptext3'),
                             'wptext3 is not matched')

        self.compare_values1(self.ms_data.get('employer4'), C2.ui_emp4, self.ams_data.get('employer4'),
                             'Employer4 is not matched')
        self.compare_values1(self.ms_data.get('industry4'), C2.ui_indus4, self.ams_data.get('industry4'),
                             'Industry4 is not matched')
        self.compare_values1(self.ms_data.get('experience4'), C2.ui_expe4, int(self.ams_data.get('experience4')),
                             'Experience4 is not matched')
        self.compare_values1(self.ms_data.get('designation4'), C2.ui_desg4, self.ams_data.get('designation4'),
                             'Designation4 is not matched')
        self.compare_values1(self.ms_data.get('wptext4'), C2.ui_wpt4, self.ams_data.get('wptext4'),
                             'wptext4 is not matched')
        self.compare_values1(self.ms_data.get('usn'), C2.ui_usn, self.ams_data.get('usn'), 'USN is not matched')
        self.compare_values1(int(self.ms_data.get('aadhar')), int(C2.ui_aadhar), int(self.ams_data.get('aadhar')),
                             'Aadhar is not matched')
        self.compare_values1(int(self.ms_data.get('currentCtc')), int(C2.ui_current_ctc),
                             int(self.ams_data.get('currentCtc')), 'currentCTC is not matched')

        self.compare_values1(self.ms_data.get('tenthBoard2'), C2.ui_tenthboard2, self.ams_data.get('tenthBoard2'),
                             'TenthBoard2 is not matched')
        self.compare_values1(self.ms_data.get('tenthCity'), C2.ui_tenthcity, self.ams_data.get('tenthCity'),
                             'TenthCity is not matched')
        self.compare_values1(self.ms_data.get('tenthState'), C2.ui_tenthstate, self.ams_data.get('tenthState'),
                             'TenthState is not matched')
        self.compare_values1(self.ms_data.get('tenthOutOf'), C2.ui_tenthoutof, self.ams_data.get('tenthOutOf'),
                             'TenthOutof is not matched')

        self.compare_values1(self.ms_data.get('twelthBoard2'), C2.ui_twelthboard2, self.ams_data.get('twelthBoard2'),
                             'TwelthBoard2 is not matched')
        self.compare_values1(self.ms_data.get('twelthCity'), C2.ui_twelthcity, self.ams_data.get('twelthCity'),
                             'TwelthCity is not matched')
        self.compare_values1(self.ms_data.get('twelthState'), C2.ui_twelthstate, self.ams_data.get('twelthState'),
                             'TwelthState is not matched')
        self.compare_values1(self.ms_data.get('twelthOutOf'), C2.ui_twelthoutof, self.ams_data.get('twelthOutOf'),
                             'TwelthOutof is not matched')

        self.compare_values1(self.ms_data.get('ugUniversity'), C2.ui_uguniversity, self.ams_data.get('ugUniversity'),
                             'UG universityis not matched')
        self.compare_values1(self.ms_data.get('ugCity'), C2.ui_ugcity, self.ams_data.get('ugCity'),
                             'UG City is not matched')
        self.compare_values1(self.ms_data.get('ugState'), C2.ui_ugstate, self.ams_data.get('ugState'),
                             'UG state is not matched')
        self.compare_values1(self.ms_data.get('ugOutOf'), C2.ui_ugoutof, self.ams_data.get('ugOutOf'),
                             'UG Outof is not matched')

        self.compare_values1(self.ms_data.get('pgUniversity'), C2.ui_pguniversity, self.ams_data.get('pgUniversity'),
                             'PG universityis not matched')
        self.compare_values1(self.ms_data.get('pgCity'), C2.ui_pgcity, self.ams_data.get('pgCity'),
                             'PG City is not matched')
        self.compare_values1(self.ms_data.get('pgState'), C2.ui_pgstate, self.ams_data.get('pgState'),
                             'PG state is not matched')
        self.compare_values1(self.ms_data.get('pgOutOf'), C2.ui_pgoutof, self.ams_data.get('pgOutOf'),
                             'PG Outof is not matched')

        self.compare_values1(self.ms_data.get('others1University'), C2.ui_others1university,
                             self.ams_data.get('others1University'),
                             'others1 universityis not matched')
        self.compare_values1(self.ms_data.get('others1City'), C2.ui_others1city, self.ams_data.get('others1City'),
                             'others1 City is not matched')
        self.compare_values1(self.ms_data.get('others1State'), C2.ui_others1state, self.ams_data.get('others1State'),
                             'others1 state is not matched')
        self.compare_values1(self.ms_data.get('others1OutOf'), C2.ui_others1outof, self.ams_data.get('others1OutOf'),
                             'others1 Outof is not matched')

        self.compare_values1(self.ms_data.get('others2University'), C2.ui_others2university,
                             self.ams_data.get('others2University'),
                             'others2 universityis not matched')
        self.compare_values1(self.ms_data.get('others2City'), C2.ui_others2city, self.ams_data.get('others2City'),
                             'others2 City is not matched')
        self.compare_values1(self.ms_data.get('others2State'), C2.ui_others2state, self.ams_data.get('others2State'),
                             'others2 state is not matched')
        self.compare_values1(self.ms_data.get('others2OutOf'), C2.ui_others2outof, self.ams_data.get('others2OutOf'),
                             'others2 Outof is not matched')

        self.compare_values1(self.ms_data.get('finalUniversity'), C2.ui_finaluniversity,
                             self.ams_data.get('finalUniversity'), 'final universityis not matched')
        self.compare_values1(self.ms_data.get('finalCity'), C2.ui_finalcity, self.ams_data.get('finalCity'),
                             'final City is not matched')
        self.compare_values1(self.ms_data.get('finalState'), C2.ui_finalstate, self.ams_data.get('finalState'),
                             'final state is not matched')
        self.compare_values1(self.ms_data.get('finalOutOf'), C2.ui_finaloutof, self.ams_data.get('finalOutOf'),
                             'final Outof is not matched')

        self.compare_values1(self.ms_data.get('experienceInYear'), C2.ui_expinyear,
                             self.ams_data.get('totalExperience'), 'Text11 is not matched')
        self.compare_values1(self.ms_data.get('text11'), C2.ui_text11, self.ams_data.get('text11'),
                             'Text11 is not matched')
        self.compare_values1(self.ms_data.get('text12'), C2.ui_text12, self.ams_data.get('text12'),
                             'Text12 is not matched')
        self.compare_values1(self.ms_data.get('text13'), C2.ui_text13, self.ams_data.get('text13'),
                             'Text13 is not matched')
        self.compare_values1(self.ms_data.get('text14'), C2.ui_text14, self.ams_data.get('text14'),
                             'Text14 is not matched')
        self.compare_values1(self.ms_data.get('text15'), C2.ui_text15, self.ams_data.get('text15'),
                             'Text15 is not matched')

        self.compare_values1(self.ms_data.get('integer11'), C2.ui_integer11, self.ams_data.get('integer11'),
                             'Integer11 is not matched')
        self.compare_values1(self.ms_data.get('integer12'), C2.ui_integer12, self.ams_data.get('integer12'),
                             'Integer12 is not matched')
        self.compare_values1(self.ms_data.get('integer13'), C2.ui_integer13, self.ams_data.get('integer13'),
                             'Integer13 is not matched')
        self.compare_values1(self.ms_data.get('integer14'), C2.ui_integer14, self.ams_data.get('integer14'),
                             'Integer14 is not matched')
        self.compare_values1(self.ms_data.get('integer15'), C2.ui_integer15, self.ams_data.get('integer15'),
                             'Integer15 is not matched')
        self.compare_values1(self.ms_data.get('country'), C2.ui_country, self.ams_data.get('country'),
                             'Country is not matched')
        self.ws.write(self.rowsize+1, 1, self.ui_ms_row_status, self.ui_ms_status_color)
        self.ws.write(self.rowsize+2, 1, self.ms_ams_row_status, self.ms_ams_status_color)
        C2.over_all_report_status(self.ui_ms_row_status, self.ms_ams_row_status)
        self.rowsize += 2
        # C2.wb_Result.save(
        #     'F:\\automation\\PythonWorkingScripts_Output\\Microsite\\Microsite_UpdateCase(' + self.current_DateTime + ').xls')
        C2.wb_Result.save(output_path_microsite_update_case)
    def over_all_report_status(self, ui_ms_row_status, ms_ams_row_status):
        pass
        if ui_ms_row_status=='Fail' or ms_ams_row_status=='Fail':
            self.over_all_status = 'Fail'
            self.over_all_status_color = self.base_style2

    def mainmethod(self):
        self.driver.get(self.url)
        # time.sleep(3)
        print ("Main Method Started")
        self.ui_usn = self.find_element_by_id('urn2', ui_input.get('usn'), "USN is not Available")
        self.ui_aadhar = self.find_element_by_id('aadhaarno', str(int(ui_input.get('aadhar'))),
                                                 "Aadhar is not Available")
        self.ui_current_ctc = self.find_element_by_id('currentctc', str(int(ui_input.get('currentCtc'))),
                                                      "current CTC is not Available")
        self.ui_fname = self.find_element_by_id('fname', ui_input.get('firstName'), "First Name is not Available")
        self.ui_mname = self.find_element_by_id('mname', ui_input.get('middleName'), "Middle Name is not Available")
        self.ui_lname = self.find_element_by_id('lname', ui_input.get('lastName'), "Last Name is not Available")
        self.ui_mobile = self.find_element_by_id('mobile', str(int(ui_input.get('mobileNumber'))),
                                                 "Mobile number is not Available")
        self.ui_altmobile = self.find_element_by_id('alternate_mobile', str(int(ui_input.get('alternateMobileNumber'))),
                                                    "Alternate Mobile is not Available")
        self.ui_pancard = self.find_element_by_id('pannumber', ui_input.get('pancard'), "PANCARD is not Available")
        self.ui_passport = self.find_element_by_id('passportnumber', ui_input.get('passport'),
                                                   "Passport is not Available")
        self.ui_gender = self.find_element_by_xpath_truefalse(".//*[@name='gender']", ui_input.get('gender'),
                                                              "Gender is not Available")
        self.ui_marital = self.find_element_by_xpath_truefalse(".//*[@name='marital']", ui_input.get('maritalStatus'),
                                                               "Marital is not Available")
        self.ui_nationality = self.find_element_by_xpath("nationality", ui_input.get('nationality'),
                                                         "Nationality is not Available")

        self.ui_tenthboard = self.find_element_by_id('tenthboard', ui_input.get('tenthBoard1'),
                                                     "Tenthboard is not Available")
        self.ui_tenthmark = self.find_element_by_id('tenthmarks', str(int(ui_input.get('tenthMark'))),
                                                    "Tenth marks is not Available")
        self.ui_tenthyop = self.find_element_by_xpath("tenthyop", str(int(ui_input.get('tenthYop'))),
                                                      "Tenth-YOP is not Available")

        self.ui_tenthboard2 = self.find_element_by_xpath("tenthboard2", ui_input.get('tenthBoard2'),
                                                         "tenthboard2 College is not Available")
        self.ui_tenthstate = self.find_element_by_xpath("tenthstate", ui_input.get('tenthState'),
                                                        "tenthstate College is not Available")

        self.ui_tenthcity = self.find_element_by_xpath("tenthcity", ui_input.get('tenthCity'),
                                                       "Tenth-CITY is not Available")
        self.ui_tenthoutof = self.find_element_by_id("tenthmarksoutof",
                                                     int(ui_input.get('tenthOutOf')), "Tenth Out of is not Available")

        self.ui_twelthboard = self.find_element_by_id('twelthboard', ui_input.get('twelthBoard1'),
                                                      "Twelthboard is not Available")
        self.ui_twelthmark = self.find_element_by_id('twelthmarks', str(int(ui_input.get('twelthMark'))),
                                                     "Twelthmark is not Available")
        self.ui_twelthyop = self.find_element_by_xpath("twelthyop", str(int(ui_input.get('twelthYop'))),
                                                       "Twelth-YOP is not Available")

        self.ui_twelthboard2 = self.find_element_by_xpath("twelthboard2", ui_input.get('twelthBoard2'),
                                                          "tenthboard2 College is not Available")

        self.ui_twelthstate = self.find_element_by_xpath("twelthstate", ui_input.get('twelthState'),
                                                         "twelth state College is not Available")
        self.ui_twelthcity = self.find_element_by_xpath("twelthcity", ui_input.get('twelthCity'),
                                                        "twelth-CITY is not Available")
        self.ui_twelthoutof = self.find_element_by_id("twelthmarksoutof",
                                                      int(ui_input.get('twelthOutOf')),
                                                      "Twelth Out of is not Available")

        self.ui_ugcollege = self.find_element_by_xpath("ugcollege", ui_input.get('ugCollege'),
                                                       "Ug college is not Available")
        self.ui_ugdegree = self.find_element_by_xpath("ugdegree", ui_input.get('ugDegree'), "ugdegree is not Available")
        self.ui_ugbranch = self.find_element_by_xpath("ugbranch", ui_input.get('ugBranch'), "ugbranch is not Available")
        self.ui_uggrade = self.find_element_by_xpath_truefalse(".//*[@name='uggrade']", 1, "uggrade is not Available")
        self.ui_ugmark = self.find_element_by_id("ugpercentage", str(int(ui_input.get('ugPercentage'))),
                                                 "ugpercentage is not Available")
        self.ui_ugyop = self.find_element_by_xpath("ugyop", str(int(ui_input.get('ugYop'))), "ugyop is not Available")

        self.ui_uguniversity = self.find_element_by_xpath("uguniversity", ui_input.get('ugUniversity'),
                                                          "Ug university is not Available")
        self.ui_ugcity = self.find_element_by_xpath("ugcity", ui_input.get('ugCity'), "ug city is not Available")
        self.ui_ugstate = self.find_element_by_xpath("ugstate", ui_input.get('ugState'), "ug state is not Available")
        self.ui_ugoutof = self.find_element_by_id("ugmarksoutof", int(ui_input.get('ugOutOf')),
                                                  "ug outof is not Available")

        self.ui_pgcollege = self.find_element_by_xpath("pgcollege", ui_input.get('pgCollege'),
                                                       "Pg college is not Available")
        self.ui_pgdegree = self.find_element_by_xpath("pgdegree", ui_input.get('pgDegree'), "pgdegree is not Available")
        self.ui_pgbranch = self.find_element_by_xpath("pgbranch", ui_input.get('pgBranch'), "pgbranch is not Available")
        self.ui_pggrade = self.find_element_by_xpath_truefalse(".//*[@name='pggrade']", 1, "pggrade is not Available")
        self.ui_pgmark = self.find_element_by_id("pgpercentage", str(int(ui_input.get('pgPercentage'))),
                                                 "pgpercentage is not Available")
        self.ui_pgyop = self.find_element_by_xpath("pgyop", str(int(ui_input.get('pgYop'))), "pgyop is not Available")

        self.ui_pguniversity = self.find_element_by_xpath("pguniversity", ui_input.get('pgUniversity'),
                                                          "PG university is not Available")
        self.ui_pgcity = self.find_element_by_xpath("pgcity", ui_input.get('pgCity'), "PG city is not Available")
        self.ui_pgstate = self.find_element_by_xpath("pgstate", ui_input.get('pgState'), "PG state is not Available")
        self.ui_pgoutof = self.find_element_by_id("pgmarksoutof", int(ui_input.get('pgOutOf')),
                                                  "PG outof is not Available")

        self.ui_others1college = self.find_element_by_xpath("hi1_college", ui_input.get('others1College'),
                                                            "Others1_college is not Available")
        self.ui_others1degree = self.find_element_by_xpath("hi1_degree", ui_input.get('others1Degree'),
                                                           "Others1_degree is not Available")
        self.ui_others1branch = self.find_element_by_xpath("hi1_branch", ui_input.get('others1Branch'),
                                                           "Others1_branch college is not Available")
        self.ui_others1grade = self.find_element_by_xpath_truefalse(".//*[@name='hi1_grade']", 1,
                                                                    "Others1_grade is not Available")
        self.ui_others1mark = self.find_element_by_id("hi1_percentage", str(int(ui_input.get('others1Percentage'))),
                                                      "Others1_percentage is not Available")
        self.ui_others1yop = self.find_element_by_xpath("hi1_yop", str(int(ui_input.get('others1Yop'))),
                                                        "Others1_yop  is not Available")

        self.ui_others1university = self.find_element_by_xpath("hi1university", ui_input.get('others1University'),
                                                               "Others1 university is not Available")
        self.ui_others1city = self.find_element_by_xpath("hi1city", ui_input.get('others1City'),
                                                         "Others1 city is not Available")
        self.ui_others1state = self.find_element_by_xpath("hi1state", ui_input.get('others1State'),
                                                          "Others1 state is not Available")
        self.ui_others1outof = self.find_element_by_id("hi1marksoutof", int(ui_input.get('others1OutOf')),
                                                       "Others1 out of is not Available")

        self.ui_others2college = self.find_element_by_xpath("hi2_college", ui_input.get('others2College'),
                                                            "Others2_college is not Available")
        self.ui_others2degree = self.find_element_by_xpath("hi2_degree", ui_input.get('others2Degree'),
                                                           "Others2_degree is not Available")
        self.ui_others2branch = self.find_element_by_xpath("hi2_branch", ui_input.get('others2Branch'),
                                                           "Others2_branch college is not Available")
        self.ui_others2grade = self.find_element_by_xpath_truefalse(".//*[@name='hi2_grade']", 1,
                                                                    "Others2_grade is not Available")
        self.ui_others2mark = self.find_element_by_id("hi2_percentage", str(int(ui_input.get('others2Percentage'))),
                                                      "Oters2_percentage is not Available")
        self.ui_others2yop = self.find_element_by_xpath("hi2_yop", str(int(ui_input.get('others2Yop'))),
                                                        "Others2_yop  is not Available")

        self.ui_others2university = self.find_element_by_xpath("hi2university", ui_input.get('others2University'),
                                                               "Others2 university is not Available")
        self.ui_others2city = self.find_element_by_xpath("hi2city", ui_input.get('others2City'),
                                                         "Others2 city is not Available")
        self.ui_others2state = self.find_element_by_xpath("hi2state", ui_input.get('others2State'),
                                                          "Others2 state is not Available")
        self.ui_others2outof = self.find_element_by_id("hi2marksoutof", int(ui_input.get('others2OutOf')),
                                                       "Others2 outof is not Available")

        self.ui_finalcollege = self.find_element_by_xpath("hi_college", ui_input.get('finalCollege'),
                                                          "Final College is not Available")
        self.ui_finaldegree = self.find_element_by_xpath("hi_degree", ui_input.get('finalDegree'),
                                                         "Final degree is not Available")
        self.ui_finalbranch = self.find_element_by_xpath("hi_branch", ui_input.get('finalBranch'),
                                                         "final branch college is not Available")
        self.ui_finalgrade = self.find_element_by_xpath_truefalse(".//*[@name='hi_grade']", 1,
                                                                  "final grade is not Available")
        self.ui_finalmark = self.find_element_by_id("hi_percentage", str(int(ui_input.get('finalPercentage'))),
                                                    "Final percentage is not Available")
        self.ui_finalyop = self.find_element_by_xpath("hi_yop", str(int(ui_input.get('finalYop'))),
                                                      "final yop  is not Available")

        self.ui_finaluniversity = self.find_element_by_xpath("hiuniversity", ui_input.get('finalUniversity'),
                                                             "Final university is not Available")
        self.ui_finalcity = self.find_element_by_xpath("hicity", ui_input.get('finalCity'),
                                                       "final city is not Available")
        self.ui_finalstate = self.find_element_by_xpath("histate", ui_input.get('finalState'),
                                                        "final state is not Available")
        self.ui_finaloutof = self.find_element_by_id("himarksoutof", int(ui_input.get('finalOutOf')),
                                                     "Final outof is not Available")

        self.ui_ca1 = self.find_element_by_id("caddressline1", ui_input.get('currentAddress1'),
                                              "Current Address1 is not Available")
        self.ui_ca2 = self.find_element_by_id("caddressline2", ui_input.get('currentAddress2'),
                                              "Current Address2 is not Available")
        self.ui_ccity = self.find_element_by_xpath("caddresscity", ui_input.get('currentCity'),
                                                   "Current City is not Available")
        self.ui_cstate = self.find_element_by_xpath("caddressstate", ui_input.get('currentState'),
                                                    "Current State College is not Available")
        self.ui_cpincode = self.find_element_by_id("caddresspincode", str(int(ui_input.get('currentPincode'))),
                                                   "Current Pincode is not Available")

        self.ui_pa1 = self.find_element_by_id("paddressline1", ui_input.get('permanentAddress1'),
                                              "Permanent Address1 is not Available")
        self.ui_pa2 = self.find_element_by_id("paddressline2", ui_input.get('permanentAddress2'),
                                              "Permanent Address2 is not Available")
        self.ui_pcity = self.find_element_by_xpath("paddresscity", ui_input.get('permanentCity'),
                                                   "Permanent City is not Available")
        self.ui_pstate = self.find_element_by_xpath("paddressstate", ui_input.get('permanentState'),
                                                    "Permanent State is not Available")
        self.ui_ppincode = self.find_element_by_id("paddresspincode", str(int(ui_input.get('permanentPincode'))),
                                                   "Permanent Pincode is not Available")
        self.ui_country = self.find_element_by_xpath("country", ui_input.get('country'),
                                                     "Country is not Available")

        self.ui_ccompany = self.find_element_by_id("cempname", ui_input.get('currentCompany'),
                                                   "Current Company is not Available")
        self.ui_cexp = self.find_element_by_id("cexp", str(int(ui_input.get('currentExperience'))),
                                               "Current Experience is not Available")
        self.ui_texp = self.find_element_by_id("texp", str(int(ui_input.get('totalExperience'))),
                                               "Total Experience is not Available")
        self.ui_bpoexp = self.find_element_by_id("expinbpo", str(int(ui_input.get('bpoExperience'))),
                                                 "BPO Experience is not Available")
        self.ui_expinyear = self.find_element_by_id("expinyear", str(int(ui_input.get('totalExperienceInYears'))),
                                                    "Experience in Year is not Available")

        self.ui_text1 = self.find_element_by_id("text1", ui_input.get('text1'), "text1 is not Available")
        self.ui_text2 = self.find_element_by_id("text2", ui_input.get('text2'), "text2 is not Available")
        self.ui_text3 = self.find_element_by_id("text3", ui_input.get('text3'), "text3 is not Available")
        self.ui_text4 = self.find_element_by_id("text4", ui_input.get('text4'), "text4 is not Available")
        self.ui_text5 = self.find_element_by_id("text5", ui_input.get('text5'), "text5 is not Available")
        self.ui_text6 = self.find_element_by_id("text6", ui_input.get('text6'), "text6 is not Available")
        self.ui_text7 = self.find_element_by_id("text7", ui_input.get('text7'), "text7  is not Available")
        self.ui_text8 = self.find_element_by_id("text8", ui_input.get('text8'), "text8  is not Available")
        self.ui_text9 = self.find_element_by_id("text9", ui_input.get('text9'), "text9  is not Available")
        self.ui_text10 = self.find_element_by_id("text10", ui_input.get('text10'), "text10 is not Available")

        self.ui_text11 = self.find_element_by_id("text11", ui_input.get('text11'), "text11 is not Available")
        self.ui_text12 = self.find_element_by_id("text12", ui_input.get('text12'), "text12 is not Available")
        self.ui_text13 = self.find_element_by_id("text13", ui_input.get('text13'), "text13 is not Available")
        self.ui_text14 = self.find_element_by_id("text14", ui_input.get('text14'), "text14 is not Available")
        self.ui_text15 = self.find_element_by_id("text15", ui_input.get('text15'), "text15 is not Available")

        self.ui_textarea1 = self.find_element_by_id("textarea1", ui_input.get('textarea1'),
                                                    "textarea1 is not Available")
        self.ui_textarea2 = self.find_element_by_id("textarea2", ui_input.get('textarea2'),
                                                    "textarea2 is not Available")
        self.ui_textarea3 = self.find_element_by_id("textarea3", ui_input.get('textarea3'),
                                                    "textarea3 is not Available")
        self.ui_textarea4 = self.find_element_by_id("textarea4", ui_input.get('textarea4'),
                                                    "textarea4 is not Available")

        self.ui_integer1 = self.find_element_by_xpath("ddown1", ui_input.get('integer1'), "Integer1 is not Available")
        self.ui_integer2 = self.find_element_by_xpath("ddown2", ui_input.get('integer2'), "Integer2 is not Available")
        self.ui_integer3 = self.find_element_by_xpath("ddown3", ui_input.get('integer3'), "Integer3 is not Available")
        self.ui_integer4 = self.find_element_by_xpath("ddown4", ui_input.get('integer4'), "Integer4 is not Available")
        self.ui_integer5 = self.find_element_by_xpath("ddown5", ui_input.get('integer5'), "Integer5 is not Available")
        self.ui_integer6 = self.find_element_by_xpath("ddown6", ui_input.get('integer6'), "Integer6 is not Available")
        self.ui_integer7 = self.find_element_by_xpath("ddown7", ui_input.get('integer7'), "Integer7 is not Available")
        self.ui_integer8 = self.find_element_by_xpath("ddown8", str(int(ui_input.get('integer8'))),
                                                      "Integer8 is not Available")
        self.ui_integer9 = self.find_element_by_xpath("ddown9", ui_input.get('integer9'), "Integer9 is not Available")
        self.ui_integer10 = self.find_element_by_xpath("ddown10", ui_input.get('integer10'),
                                                       "Integer10 is not Available")

        self.ui_integer11 = self.find_element_by_xpath("ddown11", str(int(ui_input.get('integer11'))),
                                                       "Integer11 is not Available")
        self.ui_integer12 = self.find_element_by_xpath("ddown12", str(int(ui_input.get('integer12'))),
                                                       "Integer12 is not Available")
        self.ui_integer13 = self.find_element_by_xpath("ddown13", str(int(ui_input.get('integer13'))),
                                                       "Integer13 is not Available")
        self.ui_integer14 = self.find_element_by_xpath("ddown14", str(int(ui_input.get('integer14'))),
                                                       "Integer14 is not Available")
        self.ui_integer15 = self.find_element_by_xpath("ddown15", str(int(ui_input.get('integer15'))),
                                                       "Integer15 is not Available")

        self.ui_emp1 = self.find_element_by_id("employer1", ui_input.get('employer1'), "employer1 is not Available")
        self.ui_indus1 = self.find_element_by_xpath("industry1", ui_input.get('industry1'),
                                                    "industry1 is not Available")
        self.ui_desg1 = self.find_element_by_id("designation1", ui_input.get('designation1'),
                                                "designation1 is not Available")
        self.ui_expe1 = self.find_element_by_id("experience1", int(ui_input.get('experience1')),
                                                "experience1 is not Available")
        self.ui_wpt1 = self.find_element_by_id("wp_text1", ui_input.get('wptext1'), "wp_text1 is not Available")

        self.ui_emp2 = self.find_element_by_id("employer2", ui_input.get('employer2'), "employer2 is not Available")
        self.ui_indus2 = self.find_element_by_xpath("industry2", ui_input.get('industry2'),
                                                    "industry2 is not Available")
        self.ui_desg2 = self.find_element_by_id("designation2", ui_input.get('designation2'),
                                                "designation2 is not Available")
        self.ui_expe2 = self.find_element_by_id("experience2", int(ui_input.get('experience2')),
                                                "experience2 is not Available")
        self.ui_wpt2 = self.find_element_by_id("wp_text2", ui_input.get('wptext2'), "wp_text2 is not Available")

        self.ui_emp3 = self.find_element_by_id("employer3", ui_input.get('employer3'), "employer3 is not Available")
        self.ui_indus3 = self.find_element_by_xpath("industry3", ui_input.get('industry3'),
                                                    "industry3 is not Available")
        self.ui_desg3 = self.find_element_by_id("designation3", ui_input.get('designation3'),
                                                "designation3 is not Available")
        self.ui_expe3 = self.find_element_by_id("experience3", int(ui_input.get('experience3')),
                                                "experience3 is not Available")
        self.ui_wpt3 = self.find_element_by_id("wp_text3", ui_input.get('wptext3'), "wp_text3 is not Available")

        self.ui_emp4 = self.find_element_by_id("employer4", ui_input.get('employer4'), "employer4 is not Available")
        self.ui_indus4 = self.find_element_by_xpath("industry4", ui_input.get('industry4'),
                                                    "industry4 is not Available")
        self.ui_desg4 = self.find_element_by_id("designation4", ui_input.get('designation4'),
                                                "designation4 is not Available")
        self.ui_expe4 = self.find_element_by_id("experience4", int(ui_input.get('experience4')),
                                                "experience4 is not Available")

        self.ui_email = self.find_element_by_id("email", ui_input.get('emailId').lower(), "email is not Available")
        self.ui_altmail = self.find_element_by_id("alternateemail", ui_input.get('alternateEmailId').lower(),
                                                  "alternateemail is not Available")
        self.ui_wpt4 = self.find_element_by_id("wp_text4", ui_input.get('wptext4'), "wp_text4 is not Available")
        self.declaration()
        time.sleep(1)
        self.submit()
        time.sleep(20)
        self.microsite_query()
        time.sleep(5)
        print ("Candidate ams_id is :-")
        print (self.candidate_amsid)
        if self.candidate_amsid != None:
            self.ams_query()
        else:
            print ("Candidate is not created in ams due to some technical reason")
        if self.Candidate_MS_Id != None:
            self.validation_ms_and_ui()
            self.starttime = datetime.datetime.now()
            print (self.starttime)
        else:
            print ("Candidate is not created in microsite due to some technical reason")


C2 = CreateCase()
C2.output_excel_header()
tot = len(excel_read_obj.details)
for i in range(0, tot):
    ui_input = excel_read_obj.details[i]
    C2.mainmethod()
    print ("Iteration Count:- %s" % i)
endtime = datetime.datetime.now()
endtime1 = "Ended:- %s" %endtime.strftime(" %H-%M-%S")
C2.ws.write(0, 0, "Microsite Update Case ", C2.base_style1)
C2.ws.write(0, 1, C2.over_all_status, C2.over_all_status_color)
C2.ws.write(0, 2, C2.starttime1, C2.base_style1)
C2.ws.write(0, 3, endtime1, C2.base_style1)
# C2.wb_Result.save(
#         'F:\\automation\\PythonWorkingScripts_Output\\Microsite\\'
#         'Microsite_UpdateCase(' + C2.current_DateTime + ').xls')
C2.wb_Result.save(output_path_microsite_update_case)
delete_Candidate_customes = 'delete from  candidate_customs where id = ' \
                                '(select candidatecustom_id from candidates where id=%s);' %C2.candidate_id
delete_edu_profiles = 'delete from candidate_education_profiles where candidate_id =%s;'%C2.candidate_id
delete_emp_profiles = 'delete from candidate_work_profiles where candidate_id =%s;'%C2.candidate_id
delete_candidates = 'delete from candidates where id= %s;'%C2.candidate_id
delete_microsite_candidate = 'delete from candidate where candidate_amsid= %s'%C2.candidate_id
C2.msdbconnection()
C2.cursor.execute(delete_microsite_candidate)
C2.conn1.commit()
C2.conn1.close()
C2.amsdbconnection()
C2.cursor.execute(delete_Candidate_customes)
C2.cursor.execute(delete_edu_profiles)
C2.cursor.execute(delete_emp_profiles)
C2.cursor.execute(delete_candidates)
C2.conn.commit()
C2.conn.close()
