from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.io_path import *
from selenium import webdriver
import datetime
import time
import xlwt
import mysql.connector


# from Microsite_Validation_Excel import *


class Create_Case():
    # This Script is used to Validate the Data between User Input,corresponding Microsite data and Corresponding AMS Data
    def __init__(self):
        # self.__borwser_Location = "/home/muthumurugan/Desktop/chromedriver_2.37"
        self.__url = "https://accenturetest-in.hirepro.in/automation-mandatory"
        # self.driver = webdriver.Chrome(self.__borwser_Location)
        self.driver = webdriver.Chrome(executable_path=r"F:\qa_automation\chromedriver.exe")
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")
        # file_path = 'F:\\automation\\PythonWorkingScripts_InputData\\' \
        #             'Microsite\\GenericExcelTest.xls'
        file_path = input_path_microsite_generic_case
        sheet_index = 2
        excel_read_obj.excel_read(file_path, sheet_index)

        # Below method is used for Microsite Database Connectivity

    def msdbconnection(self):
        self.conn1 = mysql.connector.connect(host='52.66.103.86',
                                             database='accenturetest052016',
                                             user='readuser',
                                             password='readuser')
        self.cursor = self.conn1.cursor()

        # Below method is used for AMS Database Connectivity

    def amsdbconnection(self):
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='accenturedb',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()

    def find_element_by_id(self, locator_id, xl_data, message):
        try:
            # time.sleep(0.3)
            if self.driver.find_element_by_id(locator_id).is_displayed() != 0:
                self.data = xl_data
                find_element = self.driver.find_element_by_id(locator_id)
                # time.sleep(0.5)
                find_element.clear()
                find_element.send_keys(self.data)

        except Exception as e:
            print (message)
            self.data = None
            # print e
        return self.data

    def find_element_by_xpath_truefalse(self, locator_id, xl_data, message):
        try:
            if self.driver.find_element_by_xpath(locator_id).is_displayed() != 0:
                self.data = xl_data
                find_element = self.driver.find_element_by_xpath(locator_id).click()

        except Exception as e:
            print (message)
            self.data = None
            # print e
        return self.data

    def find_element_by_xpath(self, locator_id, xl_data, message):
        try:
            # time.sleep(0.3)
            locator_id1 = "//select[@id='%s']" % locator_id
            if self.driver.find_element_by_xpath(locator_id1).is_displayed() == 0:
                self.data = xl_data
                input = "//select[@id='%s']/option[text()='%s']" % (locator_id, self.data)
                find_element = self.driver.find_element_by_xpath(input).click()

        except Exception as e:
            print (message)
            self.data = None
            # print e
        return self.data

    def declaration(self):
        try:
            if self.driver.find_element_by_id('declaration').is_displayed() != 0:
                self.driver.find_element_by_id('declaration').click()
                # time.sleep(0.10)
        except Exception as e:
            print ("declaration is not available")
            # print e

    def submit(self):
        try:
            if self.driver.find_element_by_xpath(".//*[@id='registerbtndiv']").is_displayed() != 0:
                self.driver.find_element_by_xpath(".//*[@id='registerbtndiv']").click()
                # time.sleep(0.10)
        except:
            print ("submit button is not available")

    # Below Method is used to get the candidate Information from Microsite Data base
    def Microsite_Query(self):

        # MS database connection Initiation done here
        C2.msdbconnection()

        self.Candidate_Id = "select candidate_id,candidate_amsid from candidate where candidate_email= '%s'" % (
            self.ui_email)
        query = self.Candidate_Id
        # print query
        # Passing Query To Method
        C2.Execute_Query(query)
        # print self.data
        # print len(self.data)

        # if result is available for the data ,then the below line will execute otherwise it will not execute
        if self.data != None:
            self.Candidate_MS_Id = self.data[0]
            self.candidate_amsid = self.data[1]
            self.MS_Cand_Personal = "select c.candidate_fname,c.candidate_mname,c.candidate_lname,c.candidate_email," \
                                    "c.candidate_email2,c.candidate_code,c.candidate_mobile,c.alternate_mobile," \
                                    "c.passport_no,c.pan_number,c.candidate_marital,cv.catalog_value_name as nationality," \
                                    "c.candidate_gender from candidate c left join catalog_values cv on " \
                                    "c.candidate_nationality= cv.ams_id  where c.candidate_id= %s;" \
                                    % (self.Candidate_MS_Id)
            query = self.MS_Cand_Personal
            # print query
            C2.Execute_Query(query)
            self.ms_fname = self.data[0]
            self.ms_mname = self.data[1]
            self.ms_lname = self.data[2]
            self.ms_email = self.data[3]
            self.ms_altmail = self.data[4]
            self.ms_urn = self.data[5]
            self.ms_mobile = self.data[6]
            self.ms_altmobile = self.data[7]
            self.ms_passport = self.data[8]
            self.ms_pancard = self.data[9]
            self.ms_marital = self.data[10]
            self.ms_nationality = self.data[11]
            self.ms_gender = self.data[12]

            self.MS_Cand_Address = "select c.current_address1,c.current_address2," \
                                   "v2.catalog_value_name as Current_City,v.catalog_value_name as Current_state," \
                                   "c.current_pincode,c.permanent_address1,c.permanent_address2," \
                                   "v3.catalog_value_name as Permanent_City," \
                                   "v1.catalog_value_name  as Permanent_state,c.permanent_pincode" \
                                   " from candidate c inner join catalog_values v on c.current_state =v.ams_id " \
                                   "inner join catalog_values v1 on c.permanent_state =v1.ams_id " \
                                   "inner join catalog_values v2 on c.current_city =v2.ams_id " \
                                   "inner join catalog_values v3 on c.permanent_city =v3.ams_id " \
                                   "where candidate_id= '%s'" % (self.Candidate_MS_Id);

            query = self.MS_Cand_Address
            # print query
            C2.Execute_Query(query)
            # print self.data
            self.ms_ca1 = self.data[0]
            self.ms_ca2 = self.data[1]
            self.ms_ccity = self.data[2]
            self.ms_cstate = self.data[3]
            self.ms_cpincode = self.data[4]
            self.ms_pa1 = self.data[5]
            self.ms_pa2 = self.data[6]
            self.ms_pcity = self.data[7]
            self.ms_pstate = self.data[8]
            self.ms_ppincode = self.data[9]

            self.MS_Cand_Work_Experience = "select current_employer,current_experience,bpo_experience," \
                                           "total_experience from candidate where candidate_id= '%s'" % (
                                               self.Candidate_MS_Id);

            query = self.MS_Cand_Work_Experience
            # print query
            C2.Execute_Query(query)
            # print self.data
            self.ms_ccompany = self.data[0]
            self.ms_cexp = self.data[1]
            self.ms_bexp = self.data[2]
            self.ms_texp = self.data[3]

            # self.MS_Education_profile = "select tenthboard,tenthyop,tenthmarks,twelthboard,twelthyop,twelthmarks,ugdegree,pgdegree,hi_degree,hi1_degree,hi2_degree,ugbranch,pgbranch,hi_branch,hi1_branch,hi2_branch,ugcgpa,pgcgpa,hi_cgpa,hi1_cgpa,hi2_cgpa,ugcollege,pgcollege,hi_college,hi1_college,hi2_college,ugyop,pgyop,hi_yop,hi1_yop,hi2_yop from candidate where candidate_id= '%s'" % ( self.Candidate_MS_Id);
            self.MS_Cand_Tenth_Details = "select c.tenthboard,c.tenthmarks,v.catalog_value_name as tenth_yop " \
                                         "from candidate c left join catalog_values v on c.tenthyop= v.ams_id " \
                                         "where candidate_id= '%s'" % (self.Candidate_MS_Id);
            query = self.MS_Cand_Tenth_Details
            # print query
            C2.Execute_Query(query)
            # print self.data
            self.ms_tenthboard = self.data[0]
            self.ms_tenthmark = self.data[1]
            self.ms_tenthyear = self.data[2]

            self.MS_Cand_Twelth_Details = " select c.twelthboard,c.twelthmarks,v.catalog_value_name as twelthyop from " \
                                          "candidate c left join catalog_values v on c.twelthyop= v.ams_id " \
                                          "where candidate_id= '%s'" % (self.Candidate_MS_Id);
            # print query
            query = self.MS_Cand_Twelth_Details
            C2.Execute_Query(query)
            # print self.data
            self.ms_twelthboard = self.data[0]
            self.ms_twelthmark = self.data[1]
            self.ms_twelthyear = self.data[2]

            self.MS_Cand_Final_Details = "select v1.catalog_value_name as hi_degree,v2.catalog_value_name as hi_branch," \
                                         "hi_cgpa,v3.catalog_value_name as hi_college,v4.catalog_value_name as hi_yop " \
                                         "from candidate c left join catalog_values v1 on c.hi_degree= v1.ams_id " \
                                         "left join catalog_values v2 on c.hi_branch= v2.ams_id " \
                                         "left join catalog_values v3 on c.hi_college= v3.ams_id " \
                                         "left join catalog_values v4 on c.hi_yop= v4.ams_id " \
                                         "where candidate_id= '%s'" % (
                                             self.Candidate_MS_Id);
            # print query
            query = self.MS_Cand_Final_Details
            C2.Execute_Query(query)
            # print self.data
            self.ms_finaldegree = self.data[0]
            self.ms_finalbranch = self.data[1]
            self.ms_finalmark = self.data[2]
            self.ms_finalcollege = self.data[3]
            self.ms_finalyop = self.data[4]

            self.MS_Cand_Custom_text = "select text1 " \
                                       "from candidate where candidate_id= '%s'" % (self.Candidate_MS_Id);
            query = self.MS_Cand_Custom_text
            # print query
            C2.Execute_Query(query)
            # print self.data
            self.ms_text1 = self.data[0]

            self.MS_Cand_Custom_TextArea = "select textarea1 from candidate where " \
                                           "candidate_id= '%s'" % (
                                               self.Candidate_MS_Id);
            query = self.MS_Cand_Custom_TextArea
            C2.Execute_Query(query)
            # print query
            # print self.data
            self.ms_textArea1 = self.data[0]

            self.MS_Cand_Custom_Integer_values = "select v1.catalog_value_name as drop_down1 from candidate c " \
                                                 "left join catalog_values v1 on c.drop_down1=v1.ams_id " \
                                                 "where candidate_id= '%s'" % (self.Candidate_MS_Id);
            query = self.MS_Cand_Custom_Integer_values
            # print query
            C2.Execute_Query(query)

            self.ms_integer1_v = self.data[0]

            self.MS_Cand_Work_Profile = "select c.employer1,c.industry1,c.designation1,c.experience1,c.wp_text1, " \
                                        "v1.catalog_value_name as industry1_v  from candidate c " \
                                        "left join catalog_values v1 on c.industry1 = v1.ams_id " \
                                        "where candidate_id= '%s'" % (self.Candidate_MS_Id);
            # self.MS_Cand_Work_Profile = "select employer1,industry1,designation1,experience1,wp_text1,employer2,industry2,designation2,experience2,wp_text2,employer3,industry3,designation3,experience3,wp_text3,employer4,industry4,designation4,experience4,wp_text4 from candidate where candidate_id= '%s'" % (
            query = self.MS_Cand_Work_Profile
            # print query
            C2.Execute_Query(query)
            # print self.data
            self.ms_emp1 = self.data[0]
            self.ms_industry1 = self.data[1]
            self.ms_designation1 = self.data[2]
            self.ms_experiance1 = self.data[3]
            self.ms_wptext1 = self.data[4]
            self.ms_industry1_v = self.data[5]


        else:
            print ("Candidate is not created in microsite")
            self.Candidate_MS_Id = None
            self.candidate_amsid = None

        # closing MS DB connection
        self.conn1.close()

    # Below Method is used to get the candidate related data from AMS Data base
    def ams_Query(self):

        # AMS DB Connection Initiation
        C2.amsdbconnection()
        self.candidate_Id = "select id from candidates where email1= '%s'" % (self.ui_email)
        # print self.candidate_Id
        query = self.candidate_Id

        # passing query to Execute Query method to execute query
        C2.Execute_Query(query)
        self.candidate_ams_id = self.data[0]
        # print self.Candidate_MS_Id
        # __Candidate_Id=
        self.ams_cand_personal = "select c.first_name,c.middle_name,c.last_name,c.email1,c.email2,c.usn,c.Mobile1," \
                                 "c.phone_office,c.passport,c.pan_card,c.marital_status,cv.value as nationality,gender " \
                                 "from candidates c left join catalog_values cv on c.nationality = cv.id " \
                                 "where c.id= '%s' " % (self.candidate_ams_id);
        query = self.ams_cand_personal
        # print query
        C2.Execute_Query(query)
        self.ams_fname = self.data[0]
        self.ams_mname = self.data[1]
        self.ams_lname = self.data[2]
        self.ams_email = self.data[3]
        self.ams_altmail = self.data[4]
        self.ams_urn = self.data[5]
        self.ams_mobile = self.data[6]
        self.ams_altmobile = self.data[7]
        self.ams_passport = self.data[8]
        self.ams_pancard = self.data[9]
        self.ams_marital = self.data[10]
        self.ams_nationality = self.data[11]
        self.ams_gender = self.data[12]
        # print self.ams_fname
        # print self.ams_mname
        # print self.ams_lname
        # print self.ams_email
        # print self.ams_altmail
        # print self.mas_urn
        # print self.ams_mobile
        # print self.ams_altmobile
        # print self.ams_passport
        # print self.ams_pancard
        # print self.ams_marital
        # print self.ams_nationality
        # print self.ams_gender

        self.ams_cand_work_experience = "select current_experience,bpo_experience,total_experience,current_employer_text " \
                                        "from candidates where id= '%s'" % (self.candidate_ams_id);
        # print self. ams_cand_work_experience
        query = self.ams_cand_work_experience
        # print query
        C2.Execute_Query(query)
        # print self.data
        self.ams_cexp = self.data[0]
        self.ams_bpoexp = self.data[1]
        self.ams_texp = self.data[2]
        self.ams_ccompany = self.data[3]
        #
        # print self.ams_ccompany
        # print self.ams_bpoexp
        # print self.ams_texp
        # print self.ams_cexp

        self.ams_cand_address = "select address1,address2 from candidates where id= '%s'" % (self.candidate_ams_id);
        query = self.ams_cand_address
        # print query
        C2.Execute_Query(query)

        # print self.data
        res = list(self.data)
        # print res
        self.current_address = [str(s).strip() for s in res[0].split(',')]
        # print self.current_address
        self.permanent_address = [str(s).strip() for s in res[1].split(',')]
        self.ams_ca1 = self.current_address[0]
        self.ams_ca2 = self.current_address[1]
        self.ams_ccity = self.current_address[2]
        self.ams_cstate = self.current_address[3]
        self.ams_cpincode = self.current_address[4]
        self.ams_pa1 = self.permanent_address[0]
        self.ams_pa2 = self.permanent_address[1]
        self.ams_pcity = self.permanent_address[2]
        self.ams_pstate = self.permanent_address[3]
        self.ams_ppincode = self.permanent_address[4]

        # print self.ams_ca1
        # print self.ams_ca2
        # print self.ams_ccity
        # print self.ams_cstate
        # print self.ams_cpincode
        # print self.ams_pa1
        # print self.ams_pa2
        # print self.ams_pcity
        # print self.ams_pstate
        # print self.ams_ppincode

        self.ams_cand_education = "select ifnull(c.degree_text,v2.value) as degree,ifnull(c.degree_type_text,v3.value) as branch," \
                                  "ifnull(c.college_text,v1.value) as college,v4.value as yop,c.percentage " \
                                  "from candidate_education_profiles c left join catalog_values v1 on c.college_id=v1.id " \
                                  "left join catalog_values v2 on c.degree_id=v2.id left join catalog_values v3 on c.degree_type_id=v3.id " \
                                  "left join catalog_values v4 on c.end_year=v4.id where candidate_id = '%s' order by yop asc" % (
                                      self.candidate_ams_id);
        # print self.ams_cand_education
        query = self.ams_cand_education
        # print query
        C2.execute_Query2(query)
        # print self.data
        self.ams_cand_tenth = [str(s).strip() for s in list(self.data[0])]
        self.ams_cand_twelth = [str(s).strip() for s in list(self.data[1])]
        self.ams_cand_final = [str(s).strip() for s in list(self.data[2])]
        # print self.ams_cand_tenth
        # print self.ams_cand_twelth
        # print self.ams_cand_ug
        # print self.ams_cand_pg
        # print self.ams_cand_others1
        # print self.ams_cand_others2
        # print self.ams_cand_final
        self.ams_cand_tenthboard = self.ams_cand_tenth[2]
        self.ams_cand_tenthyop = self.ams_cand_tenth[3]
        self.ams_cand_tenthmark = int(float(self.ams_cand_tenth[4]))
        # print self.ams_cand_tenthboard
        # print self.ams_cand_tenthyop
        # print self.ams_cand_tenthmark

        self.ams_cand_twelthboard = self.ams_cand_twelth[2]
        self.ams_cand_twelthyop = self.ams_cand_twelth[3]
        self.ams_cand_twelthmark = int(float(self.ams_cand_twelth[4]))
        # print self.ams_cand_twelthboard
        # print self.ams_cand_twelthyop
        # print self.ams_cand_twelthmark

        self.ams_cand_finaldegree = self.ams_cand_final[0]
        self.ams_cand_finalbranch = self.ams_cand_final[1]
        self.ams_cand_finalcollege = self.ams_cand_final[2]
        self.ams_cand_finalyop = self.ams_cand_final[3]
        self.ams_cand_finalmark = int(float(self.ams_cand_final[4]))
        # print self.ams_cand_finaldegree
        # print self.ams_cand_finalbranch
        # print self.ams_cand_finalcollege
        # print self.ams_cand_finalyop
        # print self.ams_cand_finalmark

        self.ams_cand_custom_textarea = "select text_area1 from candidates where id= '%s'" % (
            self.candidate_ams_id);
        # print self.ams_cand_custom_text
        query = self.ams_cand_custom_textarea
        # print query
        C2.Execute_Query(query)
        self.ams_textarea1 = self.data[0]

        self.ams_cand_custom_text = "select c.text1 from candidates c " \
                                    "where c.id= '%s'" % (self.candidate_ams_id);
        # print self.ams_cand_custom_text
        query = self.ams_cand_custom_text
        # print query
        C2.Execute_Query(query)
        # print self.data
        self.ams_text1 = self.data[0]

        self.ams_cand_custom_integer = "select v1.value as integer1 from candidates " \
                                       "c left join candidate_customs cs on c.candidatecustom_id = cs.id " \
                                       "left join catalog_values v1 on c.integer1=v1.id " \
                                       "where c.id= '%s'" % (self.candidate_ams_id);
        # print self.ams_cand_custom_integer
        query = self.ams_cand_custom_integer
        # print query
        C2.Execute_Query(query)
        self.ams_integer1 = self.data[0]

        self.ams_cand_work_profile = "select designation_text,employer_text,cv.value as industry,title,experience from candidate_work_profiles c " \
                                     "left join catalog_values cv on c.industry =cv.id where candidate_id= '%s'" % (
                                         self.candidate_ams_id);
        # print self. ams_cand_work_profile
        query = self.ams_cand_work_profile
        # print query
        C2.execute_Query2(query)
        # print "Entire data"
        # print self.data
        self.ams_cand_wp1 = [str(s).strip() for s in list(self.data[0])]

        self.ams_cand_desg1 = self.ams_cand_wp1[0]
        self.ams_cand_emp1 = self.ams_cand_wp1[1]
        self.ams_cand_indus1 = self.ams_cand_wp1[2]
        self.ams_cand_wptext1 = self.ams_cand_wp1[3]
        self.ams_cand_exp1 = self.ams_cand_wp1[4]

        # Closing AMS DB Connection
        self.conn.close()

    def Execute_Query(self, query):
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchone()
            # self.Candidate_MS_Id = data[0]
            # self.conn.commit()
        except:
            print("Hi")

    def execute_Query2(self, query):
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchall()
        except:
            print("Hi")

    def Compare_Values(self, value1, value2, mess):
        if (value1 == value2):
            self.ws.write(self.rowsize, self.a, value1, self.blackcolor_bold)
            self.a = self.a + 1
        else:
            self.ws.write(self.rowsize, self.a, value1, self.redcolor)
            self.a = self.a + 1
            self.overall_status = "Fail"
            self.overall_status_color = self.redcolor
            print (mess)

    def MS_Header(self):
        self.rowsize = 1
        self.black_color = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.blackcolor_bold = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.redcolor = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.green_bold = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Extract Resume Result')

        self.overall_status = "Pass"
        self.overall_status_color = self.green_bold
        self.ws.write(0, 0, "Microsite UI Validation", self.green_bold)

        excel_headers = ['Data', 'Status', 'Expected Message', 'Actual Message', 'First Name', 'Middle Name',
                         'Last Name', 'Email', 'Alternate Email', 'Mobile',
                         'Alternate mobile', 'Pancard', 'Passport', 'Gender', 'Marital', 'Nationality', 'Tenth Board',
                         'Tenth YOP', 'Tenth Mark', 'Twelth Board', 'Twelth YOP', 'Twelth Mark', 'Final College',
                         'Final Branch', 'Final Degree', 'Final YOP', 'Final Percentage', 'Current Address1',
                         'Current Address2', 'Current City', 'Current State', 'Current Pincode', 'Permanent Address1',
                         'Permanent Address2', 'Permanent City',
                         'Permanent State', 'Permanent Pincode', 'Current Company', 'Current Experience',
                         'Total Experience', 'Bpo Experience', 'USN', 'Text1',
                         'Integer1', 'TextArea1', 'Employer1', 'Experience1', 'Designation1', 'Industry1', 'WPTEXT1']

        total_header_column = len(excel_headers)
        for col_index in range(0, total_header_column):
            self.ws.write(self.rowsize, col_index, excel_headers[col_index], self.green_bold)

    def MS_UI_Input(self):
        self.rowsize += 1
        ui_input = [C2.ui_fname, C2.ui_mname, C2.ui_lname, C2.ui_email,
                    C2.ui_altmail, C2.ui_mobile, C2.ui_altmobile, C2.ui_pancard, C2.ui_passport,
                    C2.ui_gender, C2.ui_marital, C2.ui_nationality, C2.ui_tenthboard, C2.ui_tenthyop, C2.ui_tenthmark,
                    C2.ui_twelthboard, C2.ui_twelthyop,
                    C2.ui_twelthmark, C2.ui_finalcollege, C2.ui_finalbranch, C2.ui_finaldegree, C2.ui_finalyop,
                    C2.ui_finalmark, C2.ui_ca1, C2.ui_ca2,
                    C2.ui_ccity, C2.ui_cstate, C2.ui_cpincode, C2.ui_pa1, C2.ui_pa2, C2.ui_pcity, C2.ui_pstate,
                    C2.ui_ppincode, C2.ui_ccompany, C2.ui_cexp,
                    C2.ui_texp, C2.ui_bpoexp, C2.ui_usn, C2.ui_text1, C2.ui_integer1, C2.ui_textarea1, C2.ui_emp1,
                    C2.ui_expe1, C2.ui_desg1, C2.ui_indus1,
                    C2.ui_wpt1]

        self.ws.write(self.rowsize, 0, 'UI Input', self.black_color)
        self.ws.write(self.rowsize + 1, 0, 'MS DB Data', self.black_color)
        self.ws.write(self.rowsize + 2, 0, 'AMS DB Data', self.black_color)
        total_col_index = len(ui_input)
        for col_index in range(0, total_col_index):
            self.ws.write(self.rowsize, col_index + 4, ui_input[col_index], self.black_color)

    def validation_MS_and_UI(self):

        self.a = 4
        self.rowsize += 3
        self.Compare_Values(self.ms_fname, C2.ui_fname, 'First name is  not matched')
        self.Compare_Values(self.ms_mname, C2.ui_mname, 'Middle name is  not matched')
        self.Compare_Values(self.ms_lname, C2.ui_lname, 'Last name is not matched')
        self.Compare_Values(self.ms_email, C2.ui_email, 'Email is not matched')
        self.Compare_Values(self.ms_altmail, C2.ui_altmail, 'Alternate Email is not matched')
        self.Compare_Values(self.ms_mobile, C2.ui_mobile, 'Mobile is not matched')
        self.Compare_Values(self.ms_altmobile, C2.ui_altmobile, 'Alternate Mobile is not matched')
        self.Compare_Values(self.ms_pancard, C2.ui_pancard, 'Pancard is not matched')
        self.Compare_Values(self.ms_passport, C2.ui_passport, 'Passport is not matched')
        self.Compare_Values(self.ms_gender, C2.ui_gender, 'Gender is not matched')
        self.Compare_Values(self.ms_marital, C2.ui_marital, 'Marital is not matched')
        self.Compare_Values(self.ms_nationality, C2.ui_nationality, 'Nationality is not matched')
        self.Compare_Values(self.ms_tenthboard, C2.ui_tenthboard, 'Tenth Board is not matched')
        self.Compare_Values(self.ms_tenthyear, C2.ui_tenthyop, 'Tenth Year is not matched')
        self.Compare_Values(self.ms_tenthmark, C2.ui_tenthmark, 'Tenth Mark is not matched')
        self.Compare_Values(self.ms_twelthboard, C2.ui_twelthboard, 'Twelth Board is not matched')
        self.Compare_Values(self.ms_twelthyear, C2.ui_twelthyop, 'Twelth Year is not matched')
        self.Compare_Values(self.ms_twelthmark, C2.ui_twelthmark, 'Twelth Mark is not matched')
        self.Compare_Values(self.ms_finalcollege, C2.ui_finalcollege, 'Final College is not matched')
        self.Compare_Values(self.ms_finalbranch, C2.ui_finalbranch, 'Final Branch is not matched')
        self.Compare_Values(self.ms_finaldegree, C2.ui_finaldegree, 'final Degree is not matched')
        self.Compare_Values(self.ms_finalyop, C2.ui_finalyop, 'final YOP is not matched')
        self.Compare_Values(self.ms_finalmark, C2.ui_finalmark, 'Others2 Percentage is not matched')
        self.Compare_Values(self.ms_ca1, C2.ui_ca1, 'CurrentAddress1 is not matched')
        self.Compare_Values(self.ms_ca2, C2.ui_ca2, 'CurrentAddress2 is not matched')
        self.Compare_Values(self.ms_ccity, C2.ui_ccity, 'Current City is not matched')
        self.Compare_Values(self.ms_cstate, C2.ui_cstate, 'Current State is not matched')
        self.Compare_Values(self.ms_cpincode, C2.ui_cpincode, 'Current Pincode is not matched')
        self.Compare_Values(self.ms_pa1, C2.ui_pa1, 'Permanent Address1 is not matched')
        self.Compare_Values(self.ms_pa2, C2.ui_pa2, 'Permanent Address2 is not matched')
        self.Compare_Values(self.ms_pcity, C2.ui_pcity, 'permanent City is not matched')
        self.Compare_Values(self.ms_pstate, C2.ui_pstate, 'permanent State is not matched')
        self.Compare_Values(self.ms_ppincode, C2.ui_ppincode, 'permanent Pincode is not matched')
        self.Compare_Values(self.ms_ccompany, C2.ui_ccompany, 'Current company is not matched')
        self.Compare_Values(self.ms_cexp, C2.ui_cexp, 'Current Experience is not matched')
        self.Compare_Values(self.ms_texp, C2.ui_texp, 'Total Experience is not matched')
        self.Compare_Values(self.ms_bexp, C2.ui_bpoexp, 'BPO Experience is not matched')
        self.Compare_Values(self.ms_urn, C2.ui_usn, 'USN is not matched')
        self.Compare_Values(self.ms_text1, C2.ui_text1, 'Text1 is not matched')
        self.Compare_Values(self.ms_integer1_v, C2.ui_integer1, 'Integer1 is not matched')
        self.Compare_Values(self.ms_textArea1, C2.ui_textarea1, 'TextArea1 is not matched')
        self.Compare_Values(self.ms_emp1, C2.ui_emp1, 'Employer1 is not matched')
        self.Compare_Values(self.ms_experiance1, C2.ui_expe1, 'Experience1 is not matched')
        self.Compare_Values(self.ms_designation1, C2.ui_desg1, 'Designation1 is not matched')
        self.Compare_Values(self.ms_industry1_v, C2.ui_indus1, 'Industry1 is not matched')
        self.Compare_Values(self.ms_wptext1, C2.ui_wpt1, 'wptext1 is not matched')

        if self.candidate_amsid != None:

            self.rowsize += 4
            self.a = 4
            C2.Compare_Values(C2.ams_fname, C2.ms_fname, 'First name is  not matched')
            C2.Compare_Values(C2.ams_mname, C2.ms_mname, 'Middle name is  not matched')
            C2.Compare_Values(C2.ams_lname, C2.ms_lname, 'Last name is not matched')
            C2.Compare_Values(C2.ams_email, C2.ms_email, 'Email is not matched')
            C2.Compare_Values(C2.ams_altmail, C2.ms_altmail, 'Alternate Email is not matched')
            C2.Compare_Values(C2.ams_mobile, C2.ms_mobile, 'Mobile is not matched')
            C2.Compare_Values(C2.ams_altmobile, C2.ms_altmobile, 'Alternate Mobile is not matched')
            C2.Compare_Values(C2.ams_pancard, C2.ms_pancard, 'Pancard is not matched')
            C2.Compare_Values(C2.ams_passport, C2.ms_passport, 'Passport is not matched')
            C2.Compare_Values(C2.ams_gender, C2.ms_gender, 'Gender is not matched')
            C2.Compare_Values(C2.ams_marital, C2.ms_marital, 'Marital is not matched')
            C2.Compare_Values(C2.ams_nationality, C2.ms_nationality, 'Nationality is not matched')
            C2.Compare_Values(C2.ams_cand_tenthboard, C2.ms_tenthboard, 'Tenth Board is not matched')
            C2.Compare_Values(C2.ams_cand_tenthyop, C2.ms_tenthyear, 'Tenth Year is not matched')
            C2.Compare_Values(C2.ams_cand_tenthmark, int(C2.ms_tenthmark), 'Tenth Mark is not matched')
            C2.Compare_Values(C2.ams_cand_twelthboard, C2.ms_twelthboard, 'Twelth Board is not matched')
            C2.Compare_Values(C2.ams_cand_twelthyop, C2.ms_twelthyear, 'Twelth Year is not matched')
            C2.Compare_Values(C2.ams_cand_twelthmark, int(C2.ms_twelthmark), 'Twelth Mark is not matched')
            C2.Compare_Values(C2.ams_cand_finalcollege, C2.ms_finalcollege, 'Final College is not matched')
            C2.Compare_Values(C2.ams_cand_finalbranch, C2.ms_finalbranch, 'Final Branch is not matched')
            C2.Compare_Values(C2.ams_cand_finaldegree, C2.ms_finaldegree, 'final Degree is not matched')
            C2.Compare_Values(C2.ams_cand_finalyop, C2.ms_finalyop, 'final YOP is not matched')
            C2.Compare_Values(C2.ams_cand_finalmark, int(C2.ms_finalmark), 'Final Percentage is not matched')
            C2.Compare_Values(C2.ams_ca1, C2.ms_ca1, 'CurrentAddress1 is not matched')
            C2.Compare_Values(C2.ams_ca2, C2.ms_ca2, 'CurrentAddress2 is not matched')
            C2.Compare_Values(C2.ams_ccity, C2.ms_ccity, 'Current City is not matched')
            C2.Compare_Values(C2.ams_cstate, C2.ms_cstate, 'Current State is not matched')
            C2.Compare_Values(C2.ams_cpincode, C2.ms_cpincode, 'Current Pincode is not matched')
            C2.Compare_Values(C2.ams_pa1, C2.ms_pa1, 'Permanent Address1 is not matched')
            C2.Compare_Values(C2.ams_pa2, C2.ms_pa2, 'Permanent Address2 is not matched')
            C2.Compare_Values(C2.ams_pcity, C2.ms_pcity, 'permanent City is not matched')
            C2.Compare_Values(C2.ams_pstate, C2.ms_pstate, 'permanent State is not matched')
            C2.Compare_Values(C2.ams_ppincode, C2.ms_ppincode, 'permanent Pincode is not matched')
            C2.Compare_Values(C2.ams_ccompany, C2.ms_ccompany, 'Current company is not matched')
            C2.Compare_Values(C2.ams_cexp, int(C2.ms_cexp), 'Current Experience is not matched')
            C2.Compare_Values(C2.ams_texp, int(C2.ms_texp), 'Total Experience is not matched')
            C2.Compare_Values(C2.ams_bpoexp, int(C2.ms_bexp), 'BPO Experience is not matched')
            C2.Compare_Values(C2.ams_urn, C2.ms_urn, 'USN is not matched')
            C2.Compare_Values(C2.ams_text1, C2.ms_text1, 'Text1 is not matched')
            C2.Compare_Values(C2.ams_textarea1, C2.ms_textArea1, 'TextArea1 is not matched')
            C2.Compare_Values(C2.ams_integer1, C2.ms_integer1_v, 'Integer1 is not matched')
            C2.Compare_Values(C2.ams_cand_emp1, C2.ms_emp1, 'Employer is not matched')
            C2.Compare_Values(int(C2.ams_cand_exp1), C2.ms_experiance1, 'Experience1 is not matched')
            C2.Compare_Values(C2.ams_cand_desg1, C2.ms_designation1, 'Designation1 is not matched')
            C2.Compare_Values(C2.ams_cand_indus1, C2.ms_industry1_v, 'Industry1 is not matched')
            C2.Compare_Values(C2.ams_cand_wptext1, C2.ms_wptext1, 'wptext1 is not matched')

        else:
            print ("Candidate  is not created in ams, so we can't compare microsite and amsdb")

        C2.save()

    @staticmethod
    def convert_to_str(key_value):
        if ui_input.get(key_value):
            return str(int(ui_input.get(key_value)))
        else:
            return ui_input.get(key_value)

    @staticmethod
    def convert_to_int(key_value):
        if ui_input.get(key_value):
            return int(ui_input.get(key_value))
        else:
            return ui_input.get(key_value)

    def mainmethod(self):
        self.driver.get(self.__url)
        print ("Main Method Started")
        self.ui_usn = self.find_element_by_id('urn2', ui_input.get('usn'), "USN is not Available")
        self.ui_fname = self.find_element_by_id('fname', ui_input.get('firstName'), "First Name is not Available")
        self.ui_mname = self.find_element_by_id('mname', ui_input.get('middleName'), "Middle Name is not Available")
        self.ui_lname = self.find_element_by_id('lname', ui_input.get('lastName'), "Last Name is not Available")
        self.ui_mobile = self.find_element_by_id('mobile', C2.convert_to_str('mobile'),
                                                 "Mobile number is not Available")
        self.ui_altmobile = self.find_element_by_id('alternate_mobile', C2.convert_to_str('alternateMobile'),
                                                    "Alternate Mobile is not Available")
        self.ui_pancard = self.find_element_by_id('pannumber', ui_input.get('pancard'), "PANCARD is not Available")
        self.ui_passport = self.find_element_by_id('passportnumber', ui_input.get('passport'),
                                                   "Passport is not Available")

        self.ui_gender = self.find_element_by_xpath_truefalse(".//*[@name='gender']", C2.convert_to_int('gender'),
                                                              "Gender is not Available")

        self.ui_marital = self.find_element_by_xpath_truefalse(".//*[@name='marital']",
                                                               C2.convert_to_str('maritalStatus'),
                                                               "Marital is not Available")
        self.ui_nationality = self.find_element_by_xpath("nationality", ui_input.get('nationality'),
                                                         "Nationality is not Available")

        self.ui_tenthboard = self.find_element_by_id('tenthboard', ui_input.get('tenthBoard'),
                                                     "Tenthboard is not Available")

        self.ui_tenthmark = self.find_element_by_id('tenthmarks', C2.convert_to_str('tenthMark'), "Tenth marks is not Available")

        self.ui_tenthyop = self.find_element_by_xpath("tenthyop", C2.convert_to_str('tenthYop'),
                                                      "Tenth-YOP is not Available")

        self.ui_twelthboard = self.find_element_by_id('twelthboard', ui_input.get('twelthBoard'),
                                                      "Twelthboard is not Available")
        self.ui_twelthmark = self.find_element_by_id('twelthmarks', C2.convert_to_str('twelthMark'),
                                                     "Twelthmark is not Available")
        self.ui_twelthyop = self.find_element_by_xpath("twelthyop", C2.convert_to_str('twelthYop'),
                                                       "Twelth-YOP is not Available")

        self.ui_finalcollege = self.find_element_by_xpath("hi_college", ui_input.get('finalCollege'),
                                                          "Final College is not Available")
        self.ui_finaldegree = self.find_element_by_xpath("hi_degree", ui_input.get('finalDegree'),
                                                         "Final degree is not Available")
        self.ui_finalbranch = self.find_element_by_xpath("hi_branch", ui_input.get('finalBranch'),
                                                         "final branch is not Available")
        self.ui_finalgrade = self.find_element_by_xpath_truefalse(".//*[@name='hi_grade']", 1,
                                                                  "final grade is not Available")
        self.ui_finalmark = self.find_element_by_id("hi_percentage", C2.convert_to_str('finalPercentage'),
                                                    "Final percentage is not Available")
        self.ui_finalyop = self.find_element_by_xpath("hi_yop", C2.convert_to_str('finalYop'),
                                                      "final yop  is not Available")

        self.ui_ca1 = self.find_element_by_id("caddressline1", ui_input.get('currentAddress1'),
                                              "Current Address1 is not Available")
        self.ui_ca2 = self.find_element_by_id("caddressline2", ui_input.get('currentAddress2'),
                                              "Current Address2 is not Available")
        self.ui_ccity = self.find_element_by_xpath("caddresscity", ui_input.get('currentCity'),
                                                   "Current City is not Available")
        self.ui_cstate = self.find_element_by_xpath("caddressstate", ui_input.get('currentState'),
                                                    "Current State College is not Available")
        self.ui_cpincode = self.find_element_by_id("caddresspincode", C2.convert_to_str('currentPincode'),
                                                   "Current Pincode is not Available")

        self.ui_pa1 = self.find_element_by_id("paddressline1", ui_input.get('permanentAddress1'),
                                              "Permanent Address1 is not Available")
        self.ui_pa2 = self.find_element_by_id("paddressline2", ui_input.get('permanentAddress2'),
                                              "Permanent Address2 is not Available")
        self.ui_pcity = self.find_element_by_xpath("paddresscity", ui_input.get('permanentCity'),
                                                   "Permanent City is not Available")
        self.ui_pstate = self.find_element_by_xpath("paddressstate", ui_input.get('permanentState'),
                                                    "Permanent State is not Available")
        self.ui_ppincode = self.find_element_by_id("paddresspincode", C2.convert_to_str('permanentPincode'),
                                                   "Permanent Pincode is not Available")

        self.ui_ccompany = self.find_element_by_id("cempname", ui_input.get('currentCompany'),
                                                   "Current Company is not Available")
        self.ui_cexp = self.find_element_by_id("cexp", C2.convert_to_str('currentExperience'),
                                               "Current Experience is not Available")
        self.ui_texp = self.find_element_by_id("texp", C2.convert_to_str('totalExperience'),
                                               "Total Experience is not Available")
        self.ui_bpoexp = self.find_element_by_id("expinbpo", C2.convert_to_str('bpoExperience'),
                                                 "BPO Experience is not Available")

        self.ui_text1 = self.find_element_by_id("text1", ui_input.get('text1'), "text1 is not Available")
        self.ui_textarea1 = self.find_element_by_id("textarea1", ui_input.get('textArea1'),
                                                    "textarea1 is not Available")
        self.ui_integer1 = self.find_element_by_xpath("ddown1", ui_input.get('integer1'), "Integer1 is not Available")

        self.ui_emp1 = self.find_element_by_id("employer1", ui_input.get('employer1'), "employer1 is not Available")
        self.ui_indus1 = self.find_element_by_xpath("industry1", ui_input.get('industry1'),
                                                    "industry1 is not Available")
        self.ui_desg1 = self.find_element_by_id("designation1", ui_input.get('designation1'),
                                                "designation1 is not Available")
        self.ui_expe1 = self.find_element_by_id("experience1", C2.convert_to_str('experience1'),
                                                "experience1 is not Available")
        self.ui_wpt1 = self.find_element_by_id("wp_text1", ui_input.get('wptext1'), "wp_text1 is not Available")

        self.ui_email = self.find_element_by_id("email", ui_input.get('email'), "email is not Available")
        self.ui_altmail = self.find_element_by_id("alternateemail", ui_input.get('alternateEmail'),
                                                  "alternateemail is not Available")
        time.sleep(1)
        self.declaration()
        time.sleep(1)
        self.submit()
        time.sleep(7)
        try:
            if self.driver.find_element_by_xpath(".//*[@id='ModalWindowBody']").is_displayed() != 0:
                message = self.driver.find_element_by_xpath(".//*[@id='ModalWindowBody']").text
                C2.MS_UI_Input()
                if ui_input.get('expectedMessage') == message:
                    self.ws.write(self.rowsize, 1, "Pass", self.green_bold)
                    self.ws.write(self.rowsize, 2, ui_input.get('expectedMessage'), self.blackcolor_bold)
                    self.ws.write(self.rowsize, 3, message, self.green_bold)

                else:
                    self.ws.write(self.rowsize, 1, "Fail", self.redcolor)
                    self.ws.write(self.rowsize, 2, ui_input.get('expectedMessage'), self.blackcolor_bold)
                    self.ws.write(self.rowsize, 3, message, self.redcolor)
                    self.overall_status = "Fail"
                    self.overall_status_color = self.redcolor

                C2.save()
                self.rowsize += 3
                print ("This is model Window Block")
                print (message)

            elif self.driver.find_element_by_xpath(".//*[@id='registerbtndiv']").is_displayed() != 0:
                message = self.driver.find_element_by_id('responsediv').text
                C2.MS_UI_Input()
                if ui_input.get('expectedMessage') == message:
                    self.ws.write(self.rowsize, 1, "Pass", self.green_bold)
                    self.ws.write(self.rowsize, 2, ui_input.get('expectedMessage'), self.blackcolor_bold)
                    self.ws.write(self.rowsize, 3, message, self.green_bold)

                else:
                    self.ws.write(self.rowsize, 1, "Fail", self.redcolor)
                    self.ws.write(self.rowsize, 2, ui_input.get('expectedMessage'), self.blackcolor_bold)
                    self.ws.write(self.rowsize, 3, message, self.redcolor)
                    self.overall_status = "Fail"
                    self.overall_status_color = self.redcolor

                C2.save()
                self.rowsize += 3

                print ("This is Failure Block")
                print (message)

            else:
                self.Microsite_Query()
                # time.sleep(5)
                print("Candidate ams_id is :-")
                print (self.candidate_amsid)
                if self.candidate_amsid != None:
                    self.ams_Query()
                else:
                    print ("Candidate is not created in ams due to some technical reason")
                if self.Candidate_MS_Id != None:
                    C2.MS_UI_Input()
                    self.validation_MS_and_UI()
                    self.rowsize += 1
                else:
                    print ("Candidate is not created in microsite due to some technical reason")
        except Exception as e:
            print (e)
            print ("Candidate Id not created")

    def save(self):
        # self.wb_Result.save(
        #     'F:\\automation\\PythonWorkingScripts_Output\\Microsite\\'
        #     'UI_Functionality_VandV(' + self.__current_DateTime + ').xls')
        self.wb_Result.save(output_path_microsite_generic_case)

    def final_status(self, starttime, endtime):
        self.ws.write(0, 1, "Over All Status is :- " + self.overall_status, self.overall_status_color)
        self.ws.write(0, 2, "Total Test Case :- 53", self.green_bold)
        self.ws.write(0, 3, "Started:- " + str(starttime.strftime("%d-%m-%Y %H:%M")), self.green_bold)
        self.ws.write(0, 4, "Ended:- " + str(endtime.strftime("%d-%m-%Y %H:%M")), self.green_bold)
        C2.save()


C2 = Create_Case()
C2.MS_Header()
tot = len(excel_read_obj.details)
print (tot)
starttime = datetime.datetime.now()
print (starttime)
for i in range(0, tot):
    print ("Current Iteration Count is ", (i + 1))
    ui_input = excel_read_obj.details[i]
    C2.mainmethod()
    # print xlob.xl_first_name[i]
endtime = datetime.datetime.now()
print(endtime)
total_execution_time = endtime - starttime
C2.final_status(starttime, endtime)
print(total_execution_time)
