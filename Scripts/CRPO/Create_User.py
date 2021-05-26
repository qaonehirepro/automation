import time
import requests
import json
import xlwt
import datetime
import xlrd


class CreateUser:
    def __init__(self):

        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------
        self.header = {"content-type": "application/json"}
        self.login_request = {"LoginName": 'admin',
                              "Password": '4LWS-067',
                              "TenantAlias": "automation",
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

        # -------------------------------------------------------------------------------------------------------------
        # |                     For your reference please follow the below ID's to upload excel sheet                  |
        # -------------------------------------------------------------------------------------------------------------
        # |                |     AMS user          |    Employee Referral  |       Vendor          |       Tpo         |
        # -------------------------------------------------------------------------------------------------------------
        # | Id             |       1               |           4           |         3             |        10         |
        # -------------------------------------------------------------------------------------------------------------
        # | Roles          |   14496, 14510, 14497 |         14500         |   14496, 14510        |       14549       |
        # -------------------------------------------------------------------------------------------------------------
        # | UserBelongstoId|        NA             |           NA          |       8570            |         5         |
        # -------------------------------------------------------------------------------------------------------------
        # | Location Id's  |    25084, 25085, 25151                | Departments |      4527                           |
        # -------------------------------------------------------------------------------------------------------------

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_Typeofuser = []  # [] Initialising data from excel sheet to the variables
        self.xl_Name = []
        self.xl_Login_Name = []
        self.xl_Email = []
        self.xl_Mobile = []
        self.xl_Roles = []
        self.xl_Department = []
        self.xl_Location = []
        self.xl_enter_password = []
        self.xl_UserBelongsTo = []

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
        excelheaders = ['Comparison', 'Status', 'User Id', 'TypeofUser', 'Message', 'Name', 'Login_name', 'Email',
                        'Location', 'Mobile', 'Roles', 'Department', 'TypeofUserId', 'UserBelongs_Id']
        for headers in excelheaders:
            if headers in ['Comparison', 'Status', 'User Id', 'TypeofUser', 'Message']:
                self.ws.write(0, index, headers, self.__style2)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1
        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.user_dict = {}
        self.user_get_details = self.user_dict

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/'
                                      'PythonWorkingScripts_InputData/CRPO/User/CreateUser.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if not rows[0]:
                self.xl_Typeofuser.append(None)
            else:
                self.xl_Typeofuser.append(int(rows[0]))

            if not rows[1]:
                self.xl_Name.append(None)
            else:
                self.xl_Name.append(str(rows[1]))

            if not rows[2]:
                self.xl_Login_Name.append(None)
            else:
                self.xl_Login_Name.append(str(rows[2]))

            if not rows[3]:
                self.xl_Email.append(None)
            else:
                self.xl_Email.append(str(rows[3]))

            if not rows[4]:
                self.xl_Mobile.append(None)
            else:
                self.xl_Mobile.append(int(rows[4]))

            if not rows[6]:
                self.xl_Department.append(None)
            else:
                self.xl_Department.append(int(rows[6]))

            if not rows[7]:
                self.xl_Location.append(None)
            else:
                self.xl_Location.append(int(rows[7]))

            if not rows[8]:
                self.xl_enter_password.append(None)
            else:
                self.xl_enter_password.append(str(rows[8]))

            if not rows[9]:
                self.xl_UserBelongsTo.append(None)
            else:
                self.xl_UserBelongsTo.append(int(rows[9]))

            roles = map(int, rows[5].split(',') if isinstance(rows[5], basestring) else [rows[5]])
            self.xl_Roles.append(roles)

    def create_user(self, loop):
        # -------------------------
        # User create request
        # -------------------------
        self.create_user_request = {"UserDetails": {"Name": self.xl_Name[loop],
                                                    "UserName": self.xl_Login_Name[loop],
                                                    "Email1": self.xl_Email[loop],
                                                    "Password": self.xl_enter_password[loop],
                                                    "IsPasswordAutoGenerated": False,
                                                    "TypeOfUser": self.xl_Typeofuser[loop],
                                                    "Mobile1": self.xl_Mobile[loop],
                                                    "LocationId": self.xl_Location[loop],
                                                    "UserRoles": self.xl_Roles[loop],
                                                    "DepartmentId": self.xl_Department[loop],
                                                    "UserBelongsTo": self.xl_UserBelongsTo[loop]}
                                    }
        create_user = requests.post("https://amsin.hirepro.in/py/common/user/create_user/",
                                    headers=self.get_token,
                                    data=json.dumps(self.create_user_request, default=str), verify=False)
        create_user_response = json.loads(create_user.content)
        print create_user_response
        print create_user.content
        self.status = create_user_response['status']
        self.userId = create_user_response.get('UserId')
        self.error = create_user_response.get('error', {})
        self.message = self.error.get('errorDescription')
        if self.status == 'OK':
            print "Create User successfully"
            print "Status is", self.status
        else:
            print "user has not been created"
            print "Status is", self.status

    def user_getbyid_details(self):
        get_user_details = requests.get("https://amsin.hirepro.in/py/common/user/get_user_by_id/{}/"
                                        .format(self.userId), headers=self.get_token)
        user_details = json.loads(get_user_details.content)
        self.user_dict = user_details['UserDetails']

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
        self.ws.write(self.rowsize, 5, self.xl_Name[loop], self.__style1)
        self.ws.write(self.rowsize, 6, self.xl_Login_Name[loop], self.__style1)
        self.ws.write(self.rowsize, 7, self.xl_Email[loop], self.__style1)
        self.ws.write(self.rowsize, 8, self.xl_Location[loop], self.__style1)
        self.ws.write(self.rowsize, 9, self.xl_Mobile[loop], self.__style1)
        self.ws.write(self.rowsize, 10, ','.join(map(str, self.xl_Roles[loop])), self.__style1)
        self.ws.write(self.rowsize, 11, self.xl_Department[loop], self.__style1)
        self.ws.write(self.rowsize, 12, self.xl_Typeofuser[loop], self.__style1)
        self.ws.write(self.rowsize, 13, self.xl_UserBelongsTo[loop], self.__style1)

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
        self.ws.write(self.rowsize, 2, self.userId)
        self.ws.write(self.rowsize, 3, self.user_dict.get('TypeOfUserText'))
        if self.userId:
            self.ws.write(self.rowsize, 1, 'Pass', self.__style8)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.__style3)

        if self.message is None:
            self.ws.write(self.rowsize, 4, self.message, self.__style7)
        else:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 4, self.message, self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 4, self.message, self.__style3)
            else:
                self.ws.write(self.rowsize, 4, self.message, self.__style7)

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------

        if self.xl_Name[loop] == self.user_dict.get('Name'):
            if self.xl_Name[loop] is None:
                self.ws.write(self.rowsize, 5, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 5, self.user_dict.get('Name'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 5, self.user_dict.get('Name', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 5, self.user_dict.get('Name', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 5, 'Validation_Fail', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Login_Name[loop] == self.user_dict.get('UserName'):
            if self.xl_Login_Name[loop] is None:
                self.ws.write(self.rowsize, 6, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 6, self.user_dict.get('UserName'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 6, self.user_dict.get('UserName', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 6, self.user_dict.get('UserName', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 6, 'Validation_Fail', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Email[loop] == self.user_dict.get('Email1'):
            if self.xl_Email[loop] is None:
                self.ws.write(self.rowsize, 7, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 7, self.user_dict.get('Email1'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 7, self.user_dict.get('Email1', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 7, self.user_dict.get('Email1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 7, 'Validation_Fail', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Location[loop] == self.user_dict.get('LocationId'):
            if self.xl_Location[loop] is None:
                self.ws.write(self.rowsize, 8, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 8, self.user_dict.get('LocationId'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 8, self.user_dict.get('LocationId', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 8, self.user_dict.get('LocationId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 8, 'Validation_Fail', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Mobile[loop] == int(self.user_dict.get('Mobile1', 0) if self.user_dict.get('Mobile1') else 0):
            if self.xl_Mobile[loop] is None:
                self.ws.write(self.rowsize, 9, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 9, self.user_dict.get('Mobile1'))
        elif self.message == self.message:
            if self.message and 'User' in self.message:
                self.ws.write(self.rowsize, 9, self.user_dict.get('Mobile1', 'Duplicate'), self.__style3)
            elif self.message and 'Email already' in self.message:
                self.ws.write(self.rowsize, 9, self.user_dict.get('Mobile1', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 9, 'Validation_Fail', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Roles[loop].sort() == self.user_dict.get('UserRoles')\
                .sort() if self.user_dict.get('UserRoles') else None:
            if self.xl_Roles[loop] is None:
                self.ws.write(self.rowsize, 10,
                              self.message if self.message else 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 10, ','.join(map(str, self.user_dict.get('UserRoles'))))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 10, self.user_dict.get('UserRoles', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 10, self.user_dict.get('UserRoles', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 10, 'Validation_Fail', self.__style6)

        # if self.xl_Roles[loop].sort() == self.user_dict.get('UserRoles') \
        #         .sort() if self.user_dict.get('UserRoles') else None:
        #     self.ws.write(self.rowsize, 9, ','.join(map(str, self.user_dict.get('UserRoles'))))
        # else:
        #     self.ws.write(self.rowsize, 9, self.user_dict.get('UserRoles', 'Duplicate'), self.__style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Department[loop] == self.user_dict.get('DepartmentId'):
            if self.xl_Department[loop] is None:
                self.ws.write(self.rowsize, 11, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 11, self.user_dict.get('DepartmentId'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 11, self.user_dict.get('DepartmentId', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 11, self.user_dict.get('DepartmentId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 11, 'Validation_Fail', self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Typeofuser[loop] == self.user_dict.get('TypeOfUser'):
            if self.xl_Typeofuser[loop] is None:
                self.ws.write(self.rowsize, 12, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 12, self.user_dict.get('TypeOfUser'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 12, self.user_dict.get('TypeOfUser', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 12, self.user_dict.get('TypeOfUser', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 12, self.user_dict.get('TypeOfUser', 'Validation_Fail'), self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_UserBelongsTo[loop] == self.user_dict.get('UserBelongsToId'):
            if self.xl_UserBelongsTo[loop] is None:
                self.ws.write(self.rowsize, 13, 'Empty_Value', self.__style7)
            else:
                self.ws.write(self.rowsize, 13, self.user_dict.get('UserBelongsToId'))
        elif self.message == self.message:
            if self.message and'User' in self.message:
                self.ws.write(self.rowsize, 13, self.user_dict.get('UserBelongsToId', 'Duplicate'), self.__style3)
            elif self.message and'Email already' in self.message:
                self.ws.write(self.rowsize, 13, self.user_dict.get('UserBelongsToId', 'Duplicate'), self.__style3)
            else:
                self.ws.write(self.rowsize, 13, self.user_dict.get('UserBelongsToId', 'Validation_Fail'), self.__style6)

        # --------------------------------------------------------------------------------------------------------------
        self.rowsize += 1  # Row increment
        Obj.wb_Result.save('/home/muthumurugan/Desktop/Automation'
                           '/PythonWorkingScripts_Output/CRPO/Create USer/API_Create_User.xls')


Obj = CreateUser()
Obj.excel_data()
Total_count = len(Obj.xl_Name)
print "Number Of Rows ::", Total_count
if Obj.login == 'OK':
    for looping in range(0, Total_count):
        print "Iteration Count is ::", looping
        Obj.create_user(looping)
        if Obj.status == 'OK':
            Obj.user_getbyid_details()
        Obj.output_excel(looping)
        Obj.user_dict = {}
