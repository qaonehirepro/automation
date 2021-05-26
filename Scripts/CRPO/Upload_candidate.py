import time
import json
import requests
import xlwt
import datetime
import xlrd


class UploadCandidate:
    def __init__(self):

        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------
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

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_eventId = []  # [] Initialising data from excel sheet to the variables
        self.xl_jobRoleId = []
        self.xl_mjrId = []
        self.xl_testId = []
        self.xl_Name = []
        self.xl_FirstName = []
        self.xl_MiddleName = []
        self.xl_LastName = []
        self.xl_Mobile1 = []
        self.xl_PhoneOffice = []
        self.xl_Email1 = []
        self.xl_Email2 = []
        self.xl_Gender = []
        self.xl_MaritalStatus = []
        self.xl_DateOfBirth = []
        self.xl_USN = []
        self.xl_Address1 = []
        self.xl_Address2 = []
        self.xl_PanNo = []
        self.xl_PassportNo = []
        self.xl_CurrentLocationId = []
        self.xl_TotalExperienceInMonths = []
        self.xl_Country = []
        self.xl_HierarchyId = []
        self.xl_Nationality = []
        self.xl_Sensitivity = []
        self.xl_StatusId = []
        self.xl_FinalPercentage = []
        self.xl_FinalEndYear = []
        self.xl_FinalDegreeId = []
        self.xl_FinalCollegeId = []
        self.xl_FinalDegreeTypeId = []
        self.xl_10thDegreeId = []
        self.xl_10thPercentage = []
        self.xl_10thEndYear = []
        self.xl_12thDegreeId = []
        self.xl_12thPercentage = []
        self.xl_12thEndYear = []
        self.xl_SourceId = []
        self.xl_CampusId = []
        self.xl_SourceType = []
        self.xl_Experience = []
        self.xl_EmployerId = []
        self.xl_DesignationId = []
        self.xl_Expertise = []
        self.xl_NoticePeriod = []
        self.xl_Integer1 = []
        self.xl_Integer2 = []
        self.xl_Integer3 = []
        self.xl_Integer4 = []
        self.xl_Integer5 = []
        self.xl_Integer6 = []
        self.xl_Integer7 = []
        self.xl_Integer8 = []
        self.xl_Integer9 = []
        self.xl_Integer10 = []
        self.xl_Integer11 = []
        self.xl_Integer12 = []
        self.xl_Integer13 = []
        self.xl_Integer14 = []
        self.xl_Integer15 = []
        self.xl_Text1 = []
        self.xl_Text2 = []
        self.xl_Text3 = []
        self.xl_Text4 = []
        self.xl_Text5 = []
        self.xl_Text6 = []
        self.xl_Text7 = []
        self.xl_Text8 = []
        self.xl_Text9 = []
        self.xl_Text10 = []
        self.xl_Text11 = []
        self.xl_Text12 = []
        self.xl_Text13 = []
        self.xl_Text14 = []
        self.xl_Text15 = []
        self.xl_TextArea1 = []
        self.xl_TextArea2 = []
        self.xl_TextArea3 = []
        self.xl_TextArea4 = []

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
        self.__style8 = xlwt.easyxf('font: name Arial, color light_orange, bold on')
        self.__style9 = xlwt.easyxf('font: name Arial, color green, bold on')

        # -------------------------------------
        # Excel sheet write for Output results
        # -------------------------------------
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Candidates')
        self.rowsize = 1
        self.size = self.rowsize
        self.col = 0

        index = 0
        excelheaders = ['Comparison', 'Candidate_Created', 'Candidate Id', 'Event Id', 'Event Name', 'Job Id',
                        'Job Name', 'Applicant Id', 'Test Id', 'Test Name',
                        'Original CId', 'Message', 'Name', 'FirstName', 'MiddleName', 'LastName', 'Mobile1',
                        'PhoneOffice', 'Email1', 'Email2', 'Gender', 'MaritalStatus', 'DateOfBirth', 'USN', 'Address1',
                        'Address2', 'Final%', 'FinalEndYear', 'FinalDegree', 'FinalCollege', 'FinalDegreeType', '10th%',
                        '10thEndYear', '12th%', '12thEndYear', 'PanNo', 'PassportNo', 'CurrentLocation',
                        'TotalExperienceInMonths', 'Country', 'HierarchyId', 'Nationality', 'Sensitivity', 'StatusId',
                        'SourceId', 'CampusId', 'SourceType', 'Experience', 'EmployerId', 'DesignationId', 'Expertise',
                        'Notice Period', 'Integer1', 'Integer2', 'Integer3', 'Integer4', 'Integer5', 'Integer6',
                        'Integer7', 'Integer8', 'Integer9', 'Integer10', 'Integer11', 'Integer12', 'Integer13',
                        'Integer14', 'Integer15', 'Text1', 'Text2', 'Text3', 'Text4', 'Text5', 'Text6', 'Text7',
                        'Text8', 'Text9', 'Text10', 'Text11', 'Text12', 'Text13', 'Text14', 'Text15', 'TextArea1',
                        'TextArea2', 'TextArea3', 'TextArea4']
        for headers in excelheaders:
            if headers in ['Comparison', 'Candidate Id', 'Original CId', 'Event Id', 'Event Name', 'Job Id',
                           'Job Name', 'Applicant Id', 'Candidate_Created', 'Message', 'Test Id', 'Test Name']:
                self.ws.write(0, index, headers, self.__style2)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1

        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.personal_details_dict = {}
        self.candidate_personal_details = self.personal_details_dict
        self.source_details_dict = {}
        self.candidate_source_details = self.source_details_dict
        self.custom_details_dict = {}
        self.candidate_custom_details = self.custom_details_dict
        self.final_degree_dict = {}
        self.candidate_final_degree_dict = self.final_degree_dict
        self.tenth_dict = {}
        self.candidate_tenth_dict = self.tenth_dict
        self.twelfth_dict = {}
        self.candidate_twelfth_dict = self.twelfth_dict
        self.experience_dict = {}
        self.candidate_experience_dict = self.experience_dict
        self.app_dict = {}
        self.event_applicant_dict = self.app_dict
        self.test_dict = {}
        self.test_detail = self.test_dict

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_InputData/'
                                      'CRPO/UploadCandidate/Candidate_Upload.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            # ------------------------------
            # Event, Job, Mjr, Test details
            # ------------------------------
            if not rows[0]:
                self.xl_eventId.append(None)
            else:
                self.xl_eventId.append(int(rows[0]))

            if not rows[1]:
                self.xl_jobRoleId.append(None)
            else:
                self.xl_jobRoleId.append(int(rows[1]))

            if not rows[2]:
                self.xl_mjrId.append(None)
            else:
                self.xl_mjrId.append(int(rows[2]))

            if not rows[3]:
                self.xl_testId.append(None)
            else:
                self.xl_testId.append(int(rows[3]))

            #  ----------------------------------------------------------
            # Personal, Source, Educational, Experience, Custom details
            # ----------------------------------------------------------
            if not rows[4]:
                self.xl_Name.append(None)
            else:
                self.xl_Name.append(rows[4])

            if not rows[5]:
                self.xl_FirstName.append(None)
            else:
                self.xl_FirstName.append(rows[5])

            if not rows[6]:
                self.xl_MiddleName.append(None)
            else:
                self.xl_MiddleName.append(rows[6])

            if not rows[7]:
                self.xl_LastName.append(None)
            else:
                self.xl_LastName.append(rows[7])

            if not rows[8]:
                self.xl_Mobile1.append(None)
            else:
                self.xl_Mobile1.append(str(int(rows[8])))

            if not rows[9]:
                self.xl_PhoneOffice.append(None)
            else:
                self.xl_PhoneOffice.append(str(int(rows[9])))

            if not rows[10]:
                self.xl_Email1.append(None)
            else:
                self.xl_Email1.append(str(rows[10]))

            if not rows[11]:
                self.xl_Email2.append(None)
            else:
                self.xl_Email2.append(str(rows[11]))

            if not rows[12]:
                self.xl_Gender.append(None)
            else:
                self.xl_Gender.append(int(rows[12]))

            if not rows[13]:
                self.xl_MaritalStatus.append(None)
            else:
                self.xl_MaritalStatus.append(int(rows[13]))

            if not rows[14]:
                self.xl_DateOfBirth.append(None)
            else:
                self.xl_DateOfBirth.append(str(rows[14]))

            if not rows[15]:
                self.xl_USN.append(None)
            else:
                self.xl_USN.append(str(rows[15]))

            if not rows[16]:
                self.xl_Address1.append(None)
            else:
                self.xl_Address1.append(str(rows[16]))

            if not rows[17]:
                self.xl_Address2.append(None)
            else:
                self.xl_Address2.append(str(rows[17]))

            if not rows[18]:
                self.xl_PanNo.append(None)
            else:
                self.xl_PanNo.append(str(rows[18]))

            if not rows[19]:
                self.xl_PassportNo.append(None)
            else:
                self.xl_PassportNo.append(str(rows[19]))

            if not rows[20]:
                self.xl_CurrentLocationId.append(None)
            else:
                self.xl_CurrentLocationId.append(int(rows[20]))

            if not rows[21]:
                self.xl_TotalExperienceInMonths.append(None)
            else:
                self.xl_TotalExperienceInMonths.append(int(rows[21]))

            if not rows[22]:
                self.xl_Country.append(None)
            else:
                self.xl_Country.append(int(rows[22]))

            if not rows[23]:
                self.xl_HierarchyId.append(None)
            else:
                self.xl_HierarchyId.append(int(rows[23]))

            if not rows[24]:
                self.xl_Nationality.append(None)
            else:
                self.xl_Nationality.append(int(rows[24]))

            if not rows[25]:
                self.xl_Sensitivity.append(None)
            else:
                self.xl_Sensitivity.append(int(rows[25]))

            if not rows[26]:
                self.xl_StatusId.append(None)
            else:
                self.xl_StatusId.append(int(rows[26]))

            if not rows[27]:
                self.xl_FinalPercentage.append(None)
            else:
                self.xl_FinalPercentage.append(float(rows[27]))

            if not rows[28]:
                self.xl_FinalEndYear.append(None)
            else:
                self.xl_FinalEndYear.append(int(rows[28]))

            if not rows[29]:
                self.xl_FinalDegreeId.append(None)
            else:
                self.xl_FinalDegreeId.append(int(rows[29]))

            if not rows[30]:
                self.xl_FinalCollegeId.append(None)
            else:
                self.xl_FinalCollegeId.append(int(rows[30]))

            if not rows[31]:
                self.xl_FinalDegreeTypeId.append(None)
            else:
                self.xl_FinalDegreeTypeId.append(int(rows[31]))

            if not rows[32]:
                self.xl_10thDegreeId.append(None)
            else:
                self.xl_10thDegreeId.append(int(rows[32]))

            if not rows[33]:
                self.xl_10thPercentage.append(None)
            else:
                self.xl_10thPercentage.append(float(rows[33]))

            if not rows[34]:
                self.xl_10thEndYear.append(None)
            else:
                self.xl_10thEndYear.append(int(rows[34]))

            if not rows[35]:
                self.xl_12thDegreeId.append(None)
            else:
                self.xl_12thDegreeId.append(int(rows[35]))

            if not rows[36]:
                self.xl_12thPercentage.append(None)
            else:
                self.xl_12thPercentage.append(float(rows[36]))

            if not rows[37]:
                self.xl_12thEndYear.append(None)
            else:
                self.xl_12thEndYear.append(int(rows[37]))

            if not rows[38]:
                self.xl_SourceId.append(None)
            else:
                self.xl_SourceId.append(int(rows[38]))

            if not rows[39]:
                self.xl_CampusId.append(None)
            else:
                self.xl_CampusId.append(int(rows[39]))

            if not rows[40]:
                self.xl_SourceType.append(None)
            else:
                self.xl_SourceType.append(int(rows[40]))

            if not rows[41]:
                self.xl_Experience.append(None)
            else:
                self.xl_Experience.append(int(rows[41]))

            if not rows[42]:
                self.xl_EmployerId.append(None)
            else:
                self.xl_EmployerId.append(int(rows[42]))

            if not rows[43]:
                self.xl_DesignationId.append(None)
            else:
                self.xl_DesignationId.append(int(rows[43]))

            if not rows[44]:
                self.xl_Expertise.append(None)
            else:
                self.xl_Expertise.append(int(rows[44]))

            if not rows[45]:
                self.xl_NoticePeriod.append(None)
            else:
                self.xl_NoticePeriod.append(int(rows[45]))

            if not rows[46]:
                self.xl_Integer1.append(None)
            else:
                self.xl_Integer1.append(int(rows[46]))

            if not rows[47]:
                self.xl_Integer2.append(None)
            else:
                self.xl_Integer2.append(int(rows[47]))

            if not rows[48]:
                self.xl_Integer3.append(None)
            else:
                self.xl_Integer3.append(int(rows[48]))

            if not rows[49]:
                self.xl_Integer4.append(None)
            else:
                self.xl_Integer4.append(int(rows[49]))

            if not rows[50]:
                self.xl_Integer5.append(None)
            else:
                self.xl_Integer5.append(int(rows[50]))

            if not rows[51]:
                self.xl_Integer6.append(None)
            else:
                self.xl_Integer6.append(int(rows[51]))

            if not rows[52]:
                self.xl_Integer7.append(None)
            else:
                self.xl_Integer7.append(int(rows[52]))

            if not rows[53]:
                self.xl_Integer8.append(None)
            else:
                self.xl_Integer8.append(int(rows[53]))

            if not rows[54]:
                self.xl_Integer9.append(None)
            else:
                self.xl_Integer9.append(int(rows[54]))

            if not rows[55]:
                self.xl_Integer10.append(None)
            else:
                self.xl_Integer10.append(int(rows[55]))

            if not rows[56]:
                self.xl_Integer11.append(None)
            else:
                self.xl_Integer11.append(int(rows[56]))

            if not rows[57]:
                self.xl_Integer12.append(None)
            else:
                self.xl_Integer12.append(int(rows[57]))

            if not rows[58]:
                self.xl_Integer13.append(None)
            else:
                self.xl_Integer13.append(int(rows[58]))

            if not rows[59]:
                self.xl_Integer14.append(None)
            else:
                self.xl_Integer14.append(int(rows[59]))

            if not rows[60]:
                self.xl_Integer15.append(None)
            else:
                self.xl_Integer15.append(int(rows[60]))

            if not rows[61]:
                self.xl_Text1.append(None)
            else:
                self.xl_Text1.append(rows[61])

            if not rows[62]:
                self.xl_Text2.append(None)
            else:
                self.xl_Text2.append(rows[62])

            if not rows[63]:
                self.xl_Text3.append(None)
            else:
                self.xl_Text3.append(rows[63])

            if not rows[64]:
                self.xl_Text4.append(None)
            else:
                self.xl_Text4.append(rows[64])

            if not rows[65]:
                self.xl_Text5.append(None)
            else:
                self.xl_Text5.append(rows[65])

            if not rows[66]:
                self.xl_Text6.append(None)
            else:
                self.xl_Text6.append(rows[66])

            if not rows[67]:
                self.xl_Text7.append(None)
            else:
                self.xl_Text7.append(rows[67])

            if not rows[68]:
                self.xl_Text8.append(None)
            else:
                self.xl_Text8.append(rows[68])

            if not rows[69]:
                self.xl_Text9.append(None)
            else:
                self.xl_Text9.append(rows[69])

            if not rows[70]:
                self.xl_Text10.append(None)
            else:
                self.xl_Text10.append(rows[70])

            if not rows[71]:
                self.xl_Text11.append(None)
            else:
                self.xl_Text11.append(rows[71])

            if not rows[72]:
                self.xl_Text12.append(None)
            else:
                self.xl_Text12.append(rows[72])

            if not rows[73]:
                self.xl_Text13.append(None)
            else:
                self.xl_Text13.append(rows[73])

            if not rows[74]:
                self.xl_Text14.append(None)
            else:
                self.xl_Text14.append(rows[74])

            if not rows[75]:
                self.xl_Text15.append(None)
            else:
                self.xl_Text15.append(rows[75])

            if not rows[76]:
                self.xl_TextArea1.append(None)
            else:
                self.xl_TextArea1.append(rows[76])

            if not rows[77]:
                self.xl_TextArea2.append(None)
            else:
                self.xl_TextArea2.append(rows[77])

            if not rows[78]:
                self.xl_TextArea3.append(None)
            else:
                self.xl_TextArea3.append(rows[78])

            if not rows[79]:
                self.xl_TextArea4.append(None)
            else:
                self.xl_TextArea4.append(rows[79])

    def bulkCreateTagCandidates(self, iteration):
        # -------------------------
        # Candidate create request
        # -------------------------
        self.create_candidate_request = {"createTagCandidates": [{
            "PersonalDetails": {
                "Name": self.xl_Name[iteration],
                "FirstName": self.xl_FirstName[iteration],
                "MiddleName": self.xl_MiddleName[iteration],
                "LastName": self.xl_LastName[iteration],
                "Mobile1": self.xl_Mobile1[iteration],
                "PhoneOffice": self.xl_PhoneOffice[iteration],
                "Email1": self.xl_Email1[iteration],
                "Email2": self.xl_Email2[iteration],
                "Gender": self.xl_Gender[iteration],
                "DateOfBirth": self.xl_DateOfBirth[iteration],
                "USN": self.xl_USN[iteration],
                "MaritalStatus": self.xl_MaritalStatus[iteration],
                "Address1": self.xl_Address1[iteration],
                "Address2": self.xl_Address2[iteration],
                "PanNo": self.xl_PanNo[iteration],
                "PassportNo": self.xl_PassportNo[iteration],
                "CurrentLocationId": self.xl_CurrentLocationId[iteration],
                "TotalExperienceInMonths": self.xl_TotalExperienceInMonths[iteration],
                "Country": self.xl_Country[iteration],
                "HierarchyId": self.xl_HierarchyId[iteration],
                "Nationality": self.xl_Nationality[iteration],
                "Sensitivity": self.xl_Sensitivity[iteration],
                "StatusId": self.xl_StatusId[iteration],
                "ExpertiseId1": self.xl_Expertise[iteration]
            },
            "EducationDetails": {
                "AddedItems": [{
                    "IsPercentage": True,
                    "Percentage": self.xl_FinalPercentage[iteration],
                    "EndYear": self.xl_FinalEndYear[iteration],
                    "IsFinal": True,
                    "DegreeId": self.xl_FinalDegreeId[iteration],
                    "CollegeId": self.xl_FinalCollegeId[iteration],
                    "DegreeTypeId": self.xl_FinalDegreeTypeId[iteration]
                }, {
                    "IsPercentage": True,
                    "DegreeId": self.xl_10thDegreeId[iteration],
                    "Percentage": self.xl_10thPercentage[iteration],
                    "EndYear": self.xl_10thEndYear[iteration],
                    "IsFinal": False
                }, {
                    "IsPercentage": False,
                    "DegreeId": self.xl_12thDegreeId[iteration],
                    "Percentage": self.xl_12thPercentage[iteration],
                    "EndYear": self.xl_12thEndYear[iteration],
                    "IsFinal": False
                }]},
            "ExperienceDetails": {
                "AddedItems": [{
                    "IsLatest": True,
                    "Experience": self.xl_Experience[iteration],
                    "EmployerId": self.xl_EmployerId[iteration],
                    "DesignationId": self.xl_DesignationId[iteration]
                }]
            },
            "CustomDetails": {
                "Integer1": self.xl_Integer1[iteration],
                "Integer2": self.xl_Integer2[iteration],
                "Integer3": self.xl_Integer3[iteration],
                "Integer4": self.xl_Integer4[iteration],
                "Integer5": self.xl_Integer5[iteration],
                "Integer6": self.xl_Integer6[iteration],
                "Integer7": self.xl_Integer7[iteration],
                "Integer8": self.xl_Integer8[iteration],
                "Integer9": self.xl_Integer9[iteration],
                "Integer10": self.xl_Integer10[iteration],
                "Integer11": self.xl_Integer11[iteration],
                "Integer12": self.xl_Integer12[iteration],
                "Integer13": self.xl_Integer13[iteration],
                "Integer14": self.xl_Integer14[iteration],
                "Integer15": self.xl_Integer15[iteration],
                "Text1": self.xl_Text1[iteration],
                "Text2": self.xl_Text2[iteration],
                "Text3": self.xl_Text3[iteration],
                "Text4": self.xl_Text4[iteration],
                "Text5": self.xl_Text5[iteration],
                "Text6": self.xl_Text6[iteration],
                "Text7": self.xl_Text7[iteration],
                "Text8": self.xl_Text8[iteration],
                "Text9": self.xl_Text9[iteration],
                "Text10": self.xl_Text10[iteration],
                "Text11": self.xl_Text11[iteration],
                "Text12": self.xl_Text12[iteration],
                "Text13": self.xl_Text13[iteration],
                "Text14": self.xl_Text14[iteration],
                "Text15": self.xl_Text15[iteration],
                "TextArea1": self.xl_TextArea1[iteration],
                "TextArea2": self.xl_TextArea2[iteration],
                "TextArea3": self.xl_TextArea3[iteration],
                "TextArea4": self.xl_TextArea4[iteration]
            },
            "PreferenceDetails": {
                "NoticePeriod": self.xl_NoticePeriod[iteration]
            },
            "SourceDetails": {
                "SourceId": self.xl_SourceId[iteration],
                "CampusId": self.xl_CampusId[iteration],
                "SourceType": self.xl_SourceType[iteration]
            },
            "applicantDetail": {
                "eventId": self.xl_eventId[iteration],
                "jobRoleId": self.xl_jobRoleId[iteration],
                "mjrId": self.xl_mjrId[iteration],
                "testId": [self.xl_testId[iteration]],
                "isCreateDuplicate": True
            }
        }],
            "Sync": "True"
        }
        create_candidate = requests.post("https://amsin.hirepro.in/py/crpo/candidate/api/v1/bulkCreateTagCandidates/",
                                         headers=self.get_token,
                                         data=json.dumps(self.create_candidate_request, default=str), verify=False)
        create_candidate_response_dict = json.loads(create_candidate.content)
        # print create_candidate_response_dict
        candidate_response_data = create_candidate_response_dict['data']
        # print candidate_response_data
        # print createcandidate.headers

        # -----------------------------------------
        # API response from bulkCreateTagCandidate
        # -----------------------------------------
        for response in candidate_response_data:
            self.isCreated = response['isCreated']
            self.OrginalCID = response.get('originalCandidateId')
            self.message = response.get('duplicateCandidateMessage')
            self.CID = response.get('candidateId')
            self.candidatesavemessage = response.get('candidateSaveMessage')
            self.applicantDetails = response.get('applicantDetails')
            self.successmessage = self.applicantDetails.get('success') if self.applicantDetails else None
            # print self.successmessage
            self.listmessage = self.successmessage.get(str(self.CID)) if self.successmessage else None

            if self.listmessage == self.listmessage:
                if self.listmessage is None:
                    self.partialmessage = None
                else:
                    self.partialmessage = ', '.join(self.listmessage)
            print self.partialmessage

            if self.isCreated:  # Always Boolean is true
                print "Create Candidate :", self.isCreated
                print "candidate Id ::", self.CID
            else:
                print "Create Candidate ::", self.isCreated
                print "Message ::", self.message

    def CandidateGetbyIdDetails(self):
        get_candidate_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_by_id/{}/"
                                              .format(self.CID), headers=self.get_token)
        candidate_details = json.loads(get_candidate_details.content)
        candidate_dict = candidate_details['Candidate']
        self.personal_details_dict = candidate_dict['PersonalDetails']
        self.source_details_dict = candidate_dict['SourceDetails']
        self.custom_details_dict = candidate_dict['CustomDetails']

    def CandidateEducationalDetails(self, loop):
        get_educational_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_education_details/{}/"
                                                .format(self.CID), headers=self.get_token)
        educational_details = json.loads(get_educational_details.content)
        educational_dict = educational_details['EducationProfile']
        for edu in educational_dict:
            if edu['DegreeId'] == self.xl_FinalDegreeId[loop]:
                self.final_degree_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_FinalDegreeId[loop]), None)
            if edu['DegreeId'] == self.xl_10thDegreeId[loop]:
                self.tenth_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_10thDegreeId[loop]), None)
            if edu['DegreeId'] == self.xl_12thDegreeId[loop]:
                self.twelfth_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_12thDegreeId[loop]), None)

    def CandidateExperienceDetails(self):
        get_experience_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_experience_details/{}/"
                                               .format(self.CID), headers=self.get_token)
        experience_details = json.loads(get_experience_details.content)
        experience_dict = experience_details['WorkProfile']
        for exp in experience_dict:
            self.experience_dict = exp

    def Event_Applicants(self, loop):
        eventapplicant_request = {
            "RecruitEventId": self.xl_eventId[loop],
            "PagingCriteriaType": {
                "MaxResults": 1000,
                "PageNumber": 1
            }
        }
        eventapplicant_api = requests.post("https://amsin.hirepro.in/py/crpo/applicant/api/v1/getAllApplicants/",
                                           headers=self.get_token,
                                           data=json.dumps(eventapplicant_request, default=str), verify=False)
        applicant_dict = json.loads(eventapplicant_api.content)
        print applicant_dict
        applicant_data = applicant_dict['data']
        if applicant_data:
            for appdata in applicant_data:
                # -----------------------------------
                # Matching with created candidate Id
                # -----------------------------------
                if appdata['CandidateId'] == self.CID:
                    self.app_dict = next((item for item in applicant_data if item['CandidateId'] == self.CID), None)
                    test_details = self.app_dict['TestUserDetailType']
                    print test_details
                    if test_details:
                        for td in test_details:
                            if td['TestId'] == self.xl_testId[loop]:
                                self.test_dict = next(
                                    (item for item in test_details if item['TestId'] == self.xl_testId[loop]), None)

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
        self.ws.write(self.rowsize, 12, self.xl_Name[loop], self.__style1)
        self.ws.write(self.rowsize, 13, self.xl_FirstName[loop], self.__style1)
        self.ws.write(self.rowsize, 14, self.xl_MiddleName[loop], self.__style1)
        self.ws.write(self.rowsize, 15, self.xl_LastName[loop], self.__style1)
        self.ws.write(self.rowsize, 16, self.xl_Mobile1[loop], self.__style1)
        self.ws.write(self.rowsize, 17, self.xl_PhoneOffice[loop], self.__style1)
        self.ws.write(self.rowsize, 18, self.xl_Email1[loop], self.__style1)
        self.ws.write(self.rowsize, 19, self.xl_Email2[loop], self.__style1)
        self.ws.write(self.rowsize, 20, self.xl_Gender[loop], self.__style1)
        self.ws.write(self.rowsize, 21, self.xl_MaritalStatus[loop], self.__style1)
        self.ws.write(self.rowsize, 22, self.xl_DateOfBirth[loop], self.__style1)
        self.ws.write(self.rowsize, 23, self.xl_USN[loop], self.__style1)
        self.ws.write(self.rowsize, 24, self.xl_Address1[loop], self.__style1)
        self.ws.write(self.rowsize, 25, self.xl_Address2[loop], self.__style1)
        self.ws.write(self.rowsize, 26, self.xl_FinalPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 27, self.xl_FinalEndYear[loop], self.__style1)
        self.ws.write(self.rowsize, 28, self.xl_FinalDegreeId[loop], self.__style1)
        self.ws.write(self.rowsize, 29, self.xl_FinalCollegeId[loop], self.__style1)
        self.ws.write(self.rowsize, 30, self.xl_FinalDegreeTypeId[loop], self.__style1)
        self.ws.write(self.rowsize, 31, self.xl_10thPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 32, self.xl_10thEndYear[loop], self.__style1)
        self.ws.write(self.rowsize, 33, self.xl_12thPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 34, self.xl_12thEndYear[loop], self.__style1)
        self.ws.write(self.rowsize, 35, self.xl_PanNo[loop], self.__style1)
        self.ws.write(self.rowsize, 36, self.xl_PassportNo[loop], self.__style1)
        self.ws.write(self.rowsize, 37, self.xl_CurrentLocationId[loop], self.__style1)
        self.ws.write(self.rowsize, 38, self.xl_TotalExperienceInMonths[loop], self.__style1)
        self.ws.write(self.rowsize, 39, self.xl_Country[loop], self.__style1)
        self.ws.write(self.rowsize, 40, self.xl_HierarchyId[loop], self.__style1)
        self.ws.write(self.rowsize, 41, self.xl_Nationality[loop], self.__style1)
        self.ws.write(self.rowsize, 42, self.xl_Sensitivity[loop], self.__style1)
        self.ws.write(self.rowsize, 43, self.xl_StatusId[loop], self.__style1)
        self.ws.write(self.rowsize, 44, self.xl_SourceId[loop], self.__style1)
        self.ws.write(self.rowsize, 45, self.xl_CampusId[loop], self.__style1)
        self.ws.write(self.rowsize, 46, self.xl_SourceType[loop], self.__style1)
        self.ws.write(self.rowsize, 47, self.xl_Experience[loop], self.__style1)
        self.ws.write(self.rowsize, 48, self.xl_EmployerId[loop], self.__style1)
        self.ws.write(self.rowsize, 49, self.xl_DesignationId[loop], self.__style1)
        self.ws.write(self.rowsize, 50, self.xl_Expertise[loop], self.__style1)
        self.ws.write(self.rowsize, 51, self.xl_NoticePeriod[loop], self.__style1)
        self.ws.write(self.rowsize, 52, self.xl_Integer1[loop], self.__style1)
        self.ws.write(self.rowsize, 53, self.xl_Integer2[loop], self.__style1)
        self.ws.write(self.rowsize, 54, self.xl_Integer3[loop], self.__style1)
        self.ws.write(self.rowsize, 55, self.xl_Integer4[loop], self.__style1)
        self.ws.write(self.rowsize, 56, self.xl_Integer5[loop], self.__style1)
        self.ws.write(self.rowsize, 57, self.xl_Integer6[loop], self.__style1)
        self.ws.write(self.rowsize, 58, self.xl_Integer7[loop], self.__style1)
        self.ws.write(self.rowsize, 59, self.xl_Integer8[loop], self.__style1)
        self.ws.write(self.rowsize, 60, self.xl_Integer9[loop], self.__style1)
        self.ws.write(self.rowsize, 61, self.xl_Integer10[loop], self.__style1)
        self.ws.write(self.rowsize, 62, self.xl_Integer11[loop], self.__style1)
        self.ws.write(self.rowsize, 63, self.xl_Integer12[loop], self.__style1)
        self.ws.write(self.rowsize, 64, self.xl_Integer13[loop], self.__style1)
        self.ws.write(self.rowsize, 65, self.xl_Integer14[loop], self.__style1)
        self.ws.write(self.rowsize, 66, self.xl_Integer15[loop], self.__style1)
        self.ws.write(self.rowsize, 67, self.xl_Text1[loop], self.__style1)
        self.ws.write(self.rowsize, 68, self.xl_Text2[loop], self.__style1)
        self.ws.write(self.rowsize, 69, self.xl_Text3[loop], self.__style1)
        self.ws.write(self.rowsize, 70, self.xl_Text4[loop], self.__style1)
        self.ws.write(self.rowsize, 71, self.xl_Text5[loop], self.__style1)
        self.ws.write(self.rowsize, 72, self.xl_Text6[loop], self.__style1)
        self.ws.write(self.rowsize, 73, self.xl_Text7[loop], self.__style1)
        self.ws.write(self.rowsize, 74, self.xl_Text8[loop], self.__style1)
        self.ws.write(self.rowsize, 75, self.xl_Text9[loop], self.__style1)
        self.ws.write(self.rowsize, 76, self.xl_Text10[loop], self.__style1)
        self.ws.write(self.rowsize, 77, self.xl_Text11[loop], self.__style1)
        self.ws.write(self.rowsize, 78, self.xl_Text12[loop], self.__style1)
        self.ws.write(self.rowsize, 79, self.xl_Text13[loop], self.__style1)
        self.ws.write(self.rowsize, 80, self.xl_Text14[loop], self.__style1)
        self.ws.write(self.rowsize, 81, self.xl_Text15[loop], self.__style1)
        self.ws.write(self.rowsize, 82, self.xl_TextArea1[loop], self.__style1)
        self.ws.write(self.rowsize, 83, self.xl_TextArea2[loop], self.__style1)
        self.ws.write(self.rowsize, 84, self.xl_TextArea3[loop], self.__style1)
        self.ws.write(self.rowsize, 85, self.xl_TextArea4[loop], self.__style1)

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)

        if self.isCreated:
            self.ws.write(self.rowsize, 1, 'Pass', self.__style9)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.__style3)
        # self.ws.write(self.rowsize, 1, str(self.isCreated))

        self.ws.write(self.rowsize, 2, self.CID)
        self.ws.write(self.rowsize, 3, self.app_dict.get('EventId', None))
        self.ws.write(self.rowsize, 4, self.app_dict.get('EventName', None))
        self.ws.write(self.rowsize, 5, self.app_dict.get('JobId', None))
        self.ws.write(self.rowsize, 6, self.app_dict.get('JobName', None))
        self.ws.write(self.rowsize, 7, self.app_dict.get('ApplicantId', None))
        self.ws.write(self.rowsize, 8, self.test_dict.get('TestId', None))
        self.ws.write(self.rowsize, 9, self.test_dict.get('TestName', None))
        self.ws.write(self.rowsize, 10, self.OrginalCID, self.__style3)
        # self.ws.write(self.rowsize, 11, self.message, self.__style3)
        if self.applicantDetails == self.applicantDetails:
            if self.partialmessage is None:
                self.ws.write(self.rowsize, 11, self.message, self.__style3)
            elif 'Saved' in self.partialmessage:
                self.ws.write(self.rowsize, 11, self.partialmessage, self.__style8)

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------
        if self.xl_Name[loop] == self.personal_details_dict.get('Name'):
            if self.xl_Name[loop] is None:
                self.ws.write(self.rowsize, 12,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 12, self.personal_details_dict.get('Name'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 12, self.personal_details_dict.get('Name', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 12, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FirstName[loop] == self.personal_details_dict.get('FirstName'):
            if self.xl_FirstName[loop] is None:
                self.ws.write(self.rowsize, 13,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 13, self.personal_details_dict.get('FirstName'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 13, self.personal_details_dict.get('FirstName', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 13, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_MiddleName[loop] == self.personal_details_dict.get('MiddleName'):
            if self.xl_MiddleName[loop] is None:
                self.ws.write(self.rowsize, 14,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 14, self.personal_details_dict.get('MiddleName'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 14, self.personal_details_dict.get('MiddleName', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 14, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_LastName[loop] == self.personal_details_dict.get('LastName'):
            if self.xl_LastName[loop] is None:
                self.ws.write(self.rowsize, 15,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 15, self.personal_details_dict.get('LastName'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 15, self.personal_details_dict.get('LastName', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 15, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if str(self.xl_Mobile1[loop]) == self.personal_details_dict.get('Mobile1'):
            if self.xl_Mobile1[loop] is None:
                self.ws.write(self.rowsize, 16,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 16, self.personal_details_dict.get('Mobile1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 16, self.personal_details_dict.get('Mobile1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 16, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if str(self.xl_PhoneOffice[loop]) == self.personal_details_dict.get('PhoneOffice'):
            if self.xl_PhoneOffice[loop] is None:
                self.ws.write(self.rowsize, 17,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 17, self.personal_details_dict.get('PhoneOffice'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 17, self.personal_details_dict.get('PhoneOffice', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 17, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Email1[loop] == self.personal_details_dict.get('Email1'):
            if self.xl_Email1[loop] is None:
                self.ws.write(self.rowsize, 18,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Email1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Email1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 18, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Email2[loop] == self.personal_details_dict.get('Email2'):
            if self.xl_Email2[loop] is None:
                self.ws.write(self.rowsize, 19,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 19, self.personal_details_dict.get('Email2'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 19, self.personal_details_dict.get('Email2', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 19, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Gender[loop] == self.personal_details_dict.get('Gender'):
            if self.xl_Gender[loop] is None:
                self.ws.write(self.rowsize, 20,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 20, self.personal_details_dict.get('Gender'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 20, self.personal_details_dict.get('Gender', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 20, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_MaritalStatus[loop] == self.personal_details_dict.get('MaritalStatus'):
            if self.xl_MaritalStatus[loop] is None:
                self.ws.write(self.rowsize, 21,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 21, self.personal_details_dict.get('MaritalStatus'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 21, self.personal_details_dict.get('MaritalStatus', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 21, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_DateOfBirth[loop] == self.personal_details_dict.get('DateOfBirth'):
            if self.xl_DateOfBirth[loop] is None:
                self.ws.write(self.rowsize, 22,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 22, self.personal_details_dict.get('DateOfBirth'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 22, self.personal_details_dict.get('DateOfBirth', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 22, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_USN[loop] == self.personal_details_dict.get('USN'):
            if self.xl_USN[loop] is None:
                self.ws.write(self.rowsize, 23,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 23, self.personal_details_dict.get('USN'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 23, self.personal_details_dict.get('USN', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 23, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Address1[loop] == self.personal_details_dict.get('Address1'):
            if self.xl_Address1[loop] is None:
                self.ws.write(self.rowsize, 24,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 24, self.personal_details_dict.get('Address1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 24, self.personal_details_dict.get('Address1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 24, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Address2[loop] == self.personal_details_dict.get('Address2'):
            if self.xl_Address2[loop] is None:
                self.ws.write(self.rowsize, 25,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 25, self.personal_details_dict.get('Address2'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 25, self.personal_details_dict.get('Address2', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 25, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalPercentage[loop] == self.final_degree_dict.get('Percentage'):
            if self.xl_FinalPercentage[loop] is None:
                self.ws.write(self.rowsize, 26,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 26, self.final_degree_dict.get('Percentage'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 26, self.final_degree_dict.get('Percentage', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 26, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalEndYear[loop] == self.final_degree_dict.get('EndYear'):
            if self.xl_FinalEndYear[loop] is None:
                self.ws.write(self.rowsize, 27,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 27, self.final_degree_dict.get('EndYear'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 27, self.final_degree_dict.get('EndYear', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 27, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalDegreeId[loop] == self.final_degree_dict.get('DegreeId'):
            if self.xl_FinalDegreeId[loop] is None:
                self.ws.write(self.rowsize, 28,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 28, self.final_degree_dict.get('DegreeId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 28, self.final_degree_dict.get('DegreeId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 28, 'Validation_Failed', self.__style6)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalCollegeId[loop] == self.final_degree_dict.get('CollegeId'):
            if self.xl_FinalCollegeId[loop] is None:
                self.ws.write(self.rowsize, 29,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 29, self.final_degree_dict.get('CollegeId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 29, self.final_degree_dict.get('CollegeId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 29, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalDegreeTypeId[loop] == self.final_degree_dict.get('DegreeTypeId'):
            if self.xl_FinalDegreeTypeId[loop] is None:
                self.ws.write(self.rowsize, 30,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 30, self.final_degree_dict.get('DegreeTypeId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 30, self.final_degree_dict.get('DegreeTypeId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 30, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_10thPercentage[loop] == self.tenth_dict.get('Percentage'):
            if self.xl_10thPercentage[loop] is None:
                self.ws.write(self.rowsize, 31,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 31, self.tenth_dict.get('Percentage'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 31, self.tenth_dict.get('Percentage', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 31, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_10thEndYear[loop] == self.tenth_dict.get('EndYear'):
            if self.xl_10thEndYear[loop] is None:
                self.ws.write(self.rowsize, 32,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 32, self.tenth_dict.get('EndYear'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 32, self.tenth_dict.get('EndYear', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 32, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_12thPercentage[loop] == self.twelfth_dict.get('Percentage'):
            if self.xl_12thPercentage[loop] is None:
                self.ws.write(self.rowsize, 33,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 33, self.twelfth_dict.get('Percentage'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 33, self.twelfth_dict.get('Percentage', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 33, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_12thEndYear[loop] == self.twelfth_dict.get('EndYear'):
            if self.xl_12thEndYear[loop] is None:
                self.ws.write(self.rowsize, 34,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 34, self.twelfth_dict.get('EndYear'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 34, self.twelfth_dict.get('EndYear', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 34, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_PanNo[loop] == self.personal_details_dict.get('PanNo'):
            if self.xl_PanNo[loop] is None:
                self.ws.write(self.rowsize, 35,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 35, self.personal_details_dict.get('PanNo'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 35, self.personal_details_dict.get('PanNo', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 35, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_PassportNo[loop] == self.personal_details_dict.get('PassportNo'):
            if self.xl_PassportNo[loop] is None:
                self.ws.write(self.rowsize, 36,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 36, self.personal_details_dict.get('PassportNo'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 36, self.personal_details_dict.get('PassportNo', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 36, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_CurrentLocationId[loop] == self.personal_details_dict.get('CurrentLocationId'):
            if self.xl_CurrentLocationId[loop] is None:
                self.ws.write(self.rowsize, 37,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 37, self.personal_details_dict.get('CurrentLocationId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 37, self.personal_details_dict.get('CurrentLocationId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 37, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TotalExperienceInMonths[loop] == self.personal_details_dict.get('TotalExperienceInYears'):
            if self.xl_TotalExperienceInMonths[loop] is None:
                self.ws.write(self.rowsize, 38,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 38,
                              '{}.{}'.format(self.personal_details_dict.get('TotalExperienceInYears'),
                                             self.personal_details_dict.get('TotalExperienceInMonths')))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 38,
                              '{}.{}'.format(self.personal_details_dict.get('TotalExperienceInYears', 'Duplicate'),
                                             self.personal_details_dict.get('TotalExperienceInMonths', 'Duplicate'),
                                             '--Converting to Year(s) & Month(s)'), self.__style3)
            else:
                self.ws.write(self.rowsize, 38, 'Validation_Failed', self.__style6)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Country[loop] == self.personal_details_dict.get('Country'):
            if self.xl_Country[loop] is None:
                self.ws.write(self.rowsize, 39,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 39, self.personal_details_dict.get('Country'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 39, self.personal_details_dict.get('Country', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 39, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_HierarchyId[loop] == self.personal_details_dict.get('HierarchyId'):
            if self.xl_HierarchyId[loop] is None:
                self.ws.write(self.rowsize, 40,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 40, self.personal_details_dict.get('HierarchyId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 40, self.personal_details_dict.get('HierarchyId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 40, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Nationality[loop] == self.personal_details_dict.get('Nationality'):
            if self.xl_Nationality[loop] is None:
                self.ws.write(self.rowsize, 41,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 41, self.personal_details_dict.get('Nationality'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 41, self.personal_details_dict.get('Nationality', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 41, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Sensitivity[loop] == self.personal_details_dict.get('Sensitivity'):
            if self.xl_Sensitivity[loop] is None:
                self.ws.write(self.rowsize, 42,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 42, self.personal_details_dict.get('Sensitivity'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 42, self.personal_details_dict.get('Sensitivity', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 42, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_StatusId[loop] == self.personal_details_dict.get('StatusId'):
            if self.xl_StatusId[loop] is None:
                self.ws.write(self.rowsize, 43,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 43, self.personal_details_dict.get('StatusId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 43, self.personal_details_dict.get('StatusId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 43, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_SourceId[loop] == self.source_details_dict.get('SourceId'):
            if self.xl_SourceId[loop] is None:
                self.ws.write(self.rowsize, 44,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 44, self.source_details_dict.get('SourceId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 44, self.source_details_dict.get('SourceId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 44, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_CampusId[loop] == self.source_details_dict.get('CampusId'):
            if self.xl_CampusId[loop] is None:
                self.ws.write(self.rowsize, 45,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 45, self.source_details_dict.get('CampusId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 45, self.source_details_dict.get('CampusId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 45, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_SourceType[loop] == self.source_details_dict.get('SourceType'):
            if self.xl_SourceType[loop] is None:
                self.ws.write(self.rowsize, 46,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 46, self.source_details_dict.get('SourceType'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 46, self.source_details_dict.get('SourceType', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 46, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Experience[loop] == self.experience_dict.get('Experience'):
            if self.xl_Experience[loop] is None:
                self.ws.write(self.rowsize, 47,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 47, self.experience_dict.get('Experience'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 47, self.experience_dict.get('Experience', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 47, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_EmployerId[loop] == self.experience_dict.get('EmployerId'):
            if self.xl_EmployerId[loop] is None:
                self.ws.write(self.rowsize, 48,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 48, self.experience_dict.get('EmployerId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 48, self.experience_dict.get('EmployerId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 48, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_DesignationId[loop] == self.experience_dict.get('DesignationId'):
            if self.xl_DesignationId[loop] is None:
                self.ws.write(self.rowsize, 49,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 49, self.experience_dict.get('DesignationId'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 49, self.experience_dict.get('DesignationId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 49, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Expertise[loop] == self.personal_details_dict.get('ExpertiseId1'):
            if self.xl_Expertise[loop] is None:
                self.ws.write(self.rowsize, 50,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 50, self.personal_details_dict.get('ExpertiseId1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 50, self.personal_details_dict.get('ExpertiseId1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 50, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_NoticePeriod[loop] == self.personal_details_dict.get('NoticePeriod'):
            if self.xl_NoticePeriod[loop] is None:
                self.ws.write(self.rowsize, 51,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 51, self.personal_details_dict.get('NoticePeriod'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 51, self.personal_details_dict.get('NoticePeriod', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 51, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer1[loop] == self.custom_details_dict.get('Integer1'):
            if self.xl_Integer1[loop] is None:
                self.ws.write(self.rowsize, 52,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 52, self.custom_details_dict.get('Integer1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 52, self.custom_details_dict.get('Integer1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 52, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer2[loop] == self.custom_details_dict.get('Integer2'):
            if self.xl_Integer2[loop] is None:
                self.ws.write(self.rowsize, 53,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 53, self.custom_details_dict.get('Integer2'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 53, self.custom_details_dict.get('Integer2', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 53, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer3[loop] == self.custom_details_dict.get('Integer3'):
            if self.xl_Integer3[loop] is None:
                self.ws.write(self.rowsize, 54,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 54, self.custom_details_dict.get('Integer3'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 54, self.custom_details_dict.get('Integer3', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 54, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer4[loop] == self.custom_details_dict.get('Integer4'):
            if self.xl_Integer4[loop] is None:
                self.ws.write(self.rowsize, 55,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 55, self.custom_details_dict.get('Integer4'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 55, self.custom_details_dict.get('Integer4', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 55, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer5[loop] == self.custom_details_dict.get('Integer5'):
            if self.xl_Integer5[loop] is None:
                self.ws.write(self.rowsize, 56,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 56, self.custom_details_dict.get('Integer5'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 56, self.custom_details_dict.get('Integer5', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 56, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer6[loop] == self.custom_details_dict.get('Integer6'):
            if self.xl_Integer6[loop] is None:
                self.ws.write(self.rowsize, 57,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 57, self.custom_details_dict.get('Integer6'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 57, self.custom_details_dict.get('Integer6', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 57, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer7[loop] == self.custom_details_dict.get('Integer7'):
            if self.xl_Integer7[loop] is None:
                self.ws.write(self.rowsize, 58,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 58, self.custom_details_dict.get('Integer7'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 58, self.custom_details_dict.get('Integer7', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 58, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer8[loop] == self.custom_details_dict.get('Integer8'):
            if self.xl_Integer8[loop] is None:
                self.ws.write(self.rowsize, 59,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 59, self.custom_details_dict.get('Integer8'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 59, self.custom_details_dict.get('Integer8', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 59, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer9[loop] == self.custom_details_dict.get('Integer9'):
            if self.xl_Integer9[loop] is None:
                self.ws.write(self.rowsize, 60,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 60, self.custom_details_dict.get('Integer9'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 60, self.custom_details_dict.get('Integer9', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 60, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer10[loop] == self.custom_details_dict.get('Integer10'):
            if self.xl_Integer10[loop] is None:
                self.ws.write(self.rowsize, 61,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 61, self.custom_details_dict.get('Integer10'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 61, self.custom_details_dict.get('Integer10', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 61, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer11[loop] == self.custom_details_dict.get('Integer11'):
            if self.xl_Integer11[loop] is None:
                self.ws.write(self.rowsize, 62,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 62, self.custom_details_dict.get('Integer11'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 62, self.custom_details_dict.get('Integer11', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 62, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer12[loop] == self.custom_details_dict.get('Integer12'):
            if self.xl_Integer12[loop] is None:
                self.ws.write(self.rowsize, 63,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 63, self.custom_details_dict.get('Integer12'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 63, self.custom_details_dict.get('Integer12', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 63, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer13[loop] == self.custom_details_dict.get('Integer13'):
            if self.xl_Integer13[loop] is None:
                self.ws.write(self.rowsize, 64,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 64, self.custom_details_dict.get('Integer13'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 64, self.custom_details_dict.get('Integer13', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 64, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer14[loop] == self.custom_details_dict.get('Integer14'):
            if self.xl_Integer14[loop] is None:
                self.ws.write(self.rowsize, 65,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 65, self.custom_details_dict.get('Integer14'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 65, self.custom_details_dict.get('Integer14', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 65, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer15[loop] == self.custom_details_dict.get('Integer15'):
            if self.xl_Integer15[loop] is None:
                self.ws.write(self.rowsize, 66,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 66, self.custom_details_dict.get('Integer15'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 66, self.custom_details_dict.get('Integer15', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 66, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text1[loop] == self.custom_details_dict.get('Text1'):
            if self.xl_Text1[loop] is None:
                self.ws.write(self.rowsize, 67,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 67, self.custom_details_dict.get('Text1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 67, self.custom_details_dict.get('Text1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 67, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text2[loop] == self.custom_details_dict.get('Text2'):
            if self.xl_Text2[loop] is None:
                self.ws.write(self.rowsize, 68,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 68, self.custom_details_dict.get('Text2'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 68, self.custom_details_dict.get('Text2', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 68, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text3[loop] == self.custom_details_dict.get('Text3'):
            if self.xl_Text3[loop] is None:
                self.ws.write(self.rowsize, 69,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 69, self.custom_details_dict.get('Text3'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 69, self.custom_details_dict.get('Text3', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 69, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text4[loop] == self.custom_details_dict.get('Text4'):
            if self.xl_Text4[loop] is None:
                self.ws.write(self.rowsize, 70,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 70, self.custom_details_dict.get('Text4'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 70, self.custom_details_dict.get('Text4', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 70, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text5[loop] == self.custom_details_dict.get('Text5'):
            if self.xl_Text5[loop] is None:
                self.ws.write(self.rowsize, 71,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 71, self.custom_details_dict.get('Text5'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 71, self.custom_details_dict.get('Text5', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 71, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text6[loop] == self.custom_details_dict.get('Text6'):
            if self.xl_Text6[loop] is None:
                self.ws.write(self.rowsize, 72,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 72, self.custom_details_dict.get('Text6'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 72, self.custom_details_dict.get('Text6', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 72, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text7[loop] == self.custom_details_dict.get('Text7'):
            if self.xl_Text7[loop] is None:
                self.ws.write(self.rowsize, 73,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 73, self.custom_details_dict.get('Text7'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 73, self.custom_details_dict.get('Text7', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 73, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text8[loop] == self.custom_details_dict.get('Text8'):
            if self.xl_Text8[loop] is None:
                self.ws.write(self.rowsize, 74,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 74, self.custom_details_dict.get('Text8'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 74, self.custom_details_dict.get('Text8', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 74, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text9[loop] == self.custom_details_dict.get('Text9'):
            if self.xl_Text9[loop] is None:
                self.ws.write(self.rowsize, 75,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 75, self.custom_details_dict.get('Text9'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 75, self.custom_details_dict.get('Text9', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 75, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text10[loop] == self.custom_details_dict.get('Text10'):
            if self.xl_Text10[loop] is None:
                self.ws.write(self.rowsize, 76,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 76, self.custom_details_dict.get('Text10'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 76, self.custom_details_dict.get('Text10', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 76, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text11[loop] == self.custom_details_dict.get('Text11'):
            if self.xl_Text11[loop] is None:
                self.ws.write(self.rowsize, 77,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 77, self.custom_details_dict.get('Text11'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 77, self.custom_details_dict.get('Text11', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 77, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text12[loop] == self.custom_details_dict.get('Text12'):
            if self.xl_Text12[loop] is None:
                self.ws.write(self.rowsize, 78,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 78, self.custom_details_dict.get('Text12'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 78, self.custom_details_dict.get('Text12', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 78, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text13[loop] == self.custom_details_dict.get('Text13'):
            if self.xl_Text13[loop] is None:
                self.ws.write(self.rowsize, 79,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 79, self.custom_details_dict.get('Text13'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 79, self.custom_details_dict.get('Text13', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 79, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text14[loop] == self.custom_details_dict.get('Text14'):
            if self.xl_Text14[loop] is None:
                self.ws.write(self.rowsize, 80,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 80, self.custom_details_dict.get('Text14'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 80, self.custom_details_dict.get('Text14', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 80, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text15[loop] == self.custom_details_dict.get('Text15'):
            if self.xl_Text15[loop] is None:
                self.ws.write(self.rowsize, 81,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 81, self.custom_details_dict.get('Text15'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 81, self.custom_details_dict.get('Text15', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 81, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea1[loop] == self.custom_details_dict.get('TextArea1'):
            if self.xl_TextArea1[loop] is None:
                self.ws.write(self.rowsize, 82,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 82, self.custom_details_dict.get('TextArea1'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 82, self.custom_details_dict.get('TextArea1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 82, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea2[loop] == self.custom_details_dict.get('TextArea2'):
            if self.xl_TextArea2[loop] is None:
                self.ws.write(self.rowsize, 83,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 83, self.custom_details_dict.get('TextArea2'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 83, self.custom_details_dict.get('TextArea2', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 83, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea3[loop] == self.custom_details_dict.get('TextArea3'):
            if self.xl_TextArea3[loop] is None:
                self.ws.write(self.rowsize, 84,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 84, self.custom_details_dict.get('TextArea3'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 84, self.custom_details_dict.get('TextArea3', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 84, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea4[loop] == self.custom_details_dict.get('TextArea4'):
            if self.xl_TextArea4[loop] is None:
                self.ws.write(self.rowsize, 85,
                              self.candidatesavemessage if self.candidatesavemessage else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 85, self.custom_details_dict.get('TextArea4'))
        elif self.candidatesavemessage == self.candidatesavemessage:
            if self.candidatesavemessage is None:
                self.ws.write(self.rowsize, 85, self.custom_details_dict.get('TextArea4', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 85, 'Validation_Failed', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        self.rowsize += 1  # Row increment
        Obj.wb_Result.save('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/'
                           'CRPO/Upload Candidate/API_UploadCandidates.xls')


Obj = UploadCandidate()
Obj.excel_data()
Total_count = len(Obj.xl_Name)
print "Number Of Rows ::", Total_count
if Obj.login == 'OK':
    for looping in range(0, Total_count):
        print "Iteration Count is ::", looping
        Obj.bulkCreateTagCandidates(looping)
        if Obj.isCreated:  # Always Boolean is true, if it is not mention
            Obj.CandidateGetbyIdDetails()
            Obj.CandidateEducationalDetails(looping)
            Obj.CandidateExperienceDetails()
            Obj.Event_Applicants(looping)
        Obj.output_excel(looping)
        Obj.personal_details_dict = {}
        Obj.source_details_dict = {}
        Obj.custom_details_dict = {}
        Obj.final_degree_dict = {}
        Obj.tenth_dict = {}
        Obj.twelfth_dict = {}
        Obj.experience_dict = {}
        Obj.app_dict = {}
        Obj.test_dict = {}
