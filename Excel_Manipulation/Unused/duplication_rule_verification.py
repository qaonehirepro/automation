import requests
import json
from Excel_Manipulation.read_excel import *
import datetime
import xlwt
import time


class VerifyDuplicationRule:

    def __init__(self):
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.rowsize = 1

        # CRPO LOGIN APPLICATION
        self.header = {"content-type": "application/json"}
        self.data = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "Automation", "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=self.header,
                                 data=json.dumps(self.data), verify=False)
        self.abc = response.json()
        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.abc.get("Token")}
        # print self.headers
        self.excelWriteHeader()
        file_path = '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_InputData/' \
                    'CRPO-RPO/Duplication_Rule/Rpotest_New1.xls'
        duplicate_sheet_index = 0
        excel_read_obj.excel_read(file_path, duplicate_sheet_index)
        data = excel_read_obj.details
        tot = len(data)
        for iter in range(0, tot):
            self.current_data = data[iter]
            self.updateDuplicateRule()
            self.checkDuplicate()

    def excelWriteHeader(self):

        self.style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('EC_Verification')
        self.ws.write(0, 0, 'Name', self.style0)
        self.ws.write(0, 1, 'Fname', self.style0)
        self.ws.write(0, 2, 'Mname', self.style0)
        self.ws.write(0, 3, 'Lname', self.style0)
        self.ws.write(0, 4, 'Email Address', self.style0)
        self.ws.write(0, 5, 'Mobile', self.style0)
        self.ws.write(0, 6, 'Phone', self.style0)
        self.ws.write(0, 7, 'Marital Status', self.style0)
        self.ws.write(0, 8, 'Gender', self.style0)
        self.ws.write(0, 9, 'DOB', self.style0)
        self.ws.write(0, 10, 'PANCARD', self.style0)
        self.ws.write(0, 11, 'PASSPORT', self.style0)
        self.ws.write(0, 12, 'Aadhar', self.style0)
        self.ws.write(0, 13, 'USN', self.style0)
        self.ws.write(0, 14, 'College', self.style0)
        self.ws.write(0, 15, 'Degree', self.style0)
        self.ws.write(0, 16, 'Location', self.style0)
        self.ws.write(0, 17, 'Total Experience(in Months)', self.style0)

        self.ws.write(0, 18, 'LinkedIn', self.style0)
        self.ws.write(0, 19, 'Facebook', self.style0)
        self.ws.write(0, 20, 'Twitter', self.style0)

        self.ws.write(0, 21, 'Text1', self.style0)
        self.ws.write(0, 22, 'Text2', self.style0)
        self.ws.write(0, 23, 'Text3', self.style0)
        self.ws.write(0, 24, 'Text4', self.style0)
        self.ws.write(0, 25, 'Text5', self.style0)
        self.ws.write(0, 26, 'Text6', self.style0)
        self.ws.write(0, 27, 'Text7', self.style0)
        self.ws.write(0, 28, 'Text8', self.style0)
        self.ws.write(0, 29, 'Text9', self.style0)
        self.ws.write(0, 30, 'Text10', self.style0)
        self.ws.write(0, 31, 'Text11', self.style0)
        self.ws.write(0, 32, 'Text12', self.style0)
        self.ws.write(0, 33, 'Text13', self.style0)
        self.ws.write(0, 34, 'Text14', self.style0)
        self.ws.write(0, 35, 'Text15', self.style0)

        self.ws.write(0, 36, 'Integer1', self.style0)
        self.ws.write(0, 37, 'Integer2', self.style0)
        self.ws.write(0, 38, 'Integer3', self.style0)
        self.ws.write(0, 39, 'Integer4', self.style0)
        self.ws.write(0, 40, 'Integer5', self.style0)
        self.ws.write(0, 41, 'Integer6', self.style0)
        self.ws.write(0, 42, 'Integer7', self.style0)
        self.ws.write(0, 43, 'Integer8', self.style0)
        self.ws.write(0, 44, 'Integer9', self.style0)
        self.ws.write(0, 45, 'Integer10', self.style0)
        self.ws.write(0, 46, 'Integer11', self.style0)
        self.ws.write(0, 47, 'Integer12', self.style0)
        self.ws.write(0, 48, 'Integer13', self.style0)
        self.ws.write(0, 49, 'Integer14', self.style0)
        self.ws.write(0, 50, 'Integer15', self.style0)
        self.ws.write(0, 51, 'Duplicate Rule', self.style0)

        self.ws.write(0, 52, 'Expected Status', self.style0)
        self.ws.write(0, 53, 'Actual Status', self.style0)
        self.ws.write(0, 54, 'Expected Message', self.style0)
        self.ws.write(0, 55, 'Actual Message', self.style0)
        self.ws.write(0, 56, 'Status', self.style0)
        # self.ws.write(0, 55, 'Message', self.__style0)

    def updateDuplicateRule(self):
        self.update_json_data = self.current_data.get('DuplicationRuleJson')
        # print self.update_json_data
        self.data1 = {"AppPreference": {"Id": 3595, "Content": self.update_json_data,
                                        "Type": "duplication_conf.default"}, "IsTenantGlobal": "true"}
        r = requests.post("https://amsin.hirepro.in/py/common/common_app_utils/save_app_preferences/",
                          headers=self.headers, data=json.dumps(self.data1, default=str), verify=False)
        # print r.content
        # print r.status_code

    def checkDuplicate(self):
        convert_date_of_birth = self.current_data.get('DateOfBirth')
        self.date_of_birth = datetime.datetime(
            *xlrd.xldate_as_tuple(convert_date_of_birth, excel_read_obj.excel_file.datemode))
        self.date_of_birth = self.date_of_birth.strftime("%Y-%m-%d")

        self.data = {"FirstName": self.current_data.get('FirstName'), "MiddleName": self.current_data.get('MiddleName'),
                     "LastName": self.current_data.get('LastName'),
                     "Email1": self.current_data.get('EmailAddress'),
                     "Mobile1": int(self.current_data.get('MobileNumber')) if self.current_data.get(
                         'MobileNumber') else None,
                     "PhoneOffice": int(self.current_data.get('PhoneNumber')) if self.current_data.get(
                         'PhoneNumber') else None,
                     "MaritalStatus": int(self.current_data.get('MaritalStatus')) if self.current_data.get(
                         'MaritalStatus') else None, "Gender": int(self.current_data.get('Gender')),
                     "DateOfBirth": self.date_of_birth,
                     "PanNo": self.current_data.get('Pancard'), "PassportNo": self.current_data.get('Passport'),
                     "AadhaarNo": int(self.current_data.get('Aadhar')) if self.current_data.get('Aadhar') else None,
                     "CollegeId": int(self.current_data.get('College')) if self.current_data.get('College') else None,
                     "DegreeId": int(self.current_data.get('Degree')) if self.current_data.get('Degree') else None,
                     "USN": self.current_data.get('USN'),
                     "CurrentLocationId": int(self.current_data.get('Location')) if self.current_data.get(
                         'Location') else None,
                     "TotalExperience": int(self.current_data.get('TotalExperienceInMonths')) if self.current_data.get(
                         'TotalExperienceInMonths') else None,
                     "FacebookLink": self.current_data.get('Facebook'),
                     "TwitterLink": self.current_data.get('Twitter'),
                     "LinkedInLink": self.current_data.get('LinkedIn'),
                     "Integer1": int(self.current_data.get('Integer1')) if self.current_data.get('Integer1') else None,
                     "Integer2": int(self.current_data.get('Integer2')) if self.current_data.get('Integer2') else None,
                     "Integer3": int(self.current_data.get('Integer3')) if self.current_data.get('Integer3') else None,
                     "Integer4": int(self.current_data.get('Integer4')) if self.current_data.get('Integer4') else None,
                     "Integer5": int(self.current_data.get('Integer5')) if self.current_data.get('Integer5') else None,
                     "Integer6": int(self.current_data.get('Integer6')) if self.current_data.get('Integer6') else None,
                     "Integer7": int(self.current_data.get('Integer7')) if self.current_data.get('Integer7') else None,
                     "Integer8": int(self.current_data.get('Integer8')) if self.current_data.get('Integer8') else None,
                     "Integer9": int(self.current_data.get('Integer9')) if self.current_data.get('Integer9') else None,
                     "Integer10": int(self.current_data.get('Integer10')) if self.current_data.get(
                         'Integer10') else None,
                     "Integer11": int(self.current_data.get('Integer11')) if self.current_data.get(
                         'Integer11') else None,
                     "Integer12": int(self.current_data.get('Integer12')) if self.current_data.get(
                         'Integer12') else None,
                     "Integer13": int(self.current_data.get('Integer13')) if self.current_data.get(
                         'Integer13') else None,
                     "Integer14": int(self.current_data.get('Integer14')) if self.current_data.get(
                         'Integer14') else None,
                     "Integer15": int(self.current_data.get('Integer15')) if self.current_data.get(
                         'Integer15') else None,
                     "Text1": self.current_data.get('Text1'), "Text2": self.current_data.get('Text2'),
                     "Text3": self.current_data.get('Text3'),
                     "Text4": self.current_data.get('Text4'), "Text5": self.current_data.get('Text5'),
                     "Text6": self.current_data.get('Text6'),
                     "Text7": self.current_data.get('Text7'), "Text8": self.current_data.get('Text8'),
                     "Text9": self.current_data.get('Text9'),
                     "Text10": self.current_data.get('Text10'), "Text11": self.current_data.get('Text11'),
                     "Text12": self.current_data.get('Text12'),
                     "Text13": self.current_data.get('Text13'), "Text14": self.current_data.get('Text14'),
                     "Text15": self.current_data.get('Text15')
                     }
        # print self.data

        r = requests.post("https://amsin.hirepro.in/py/rpo/candidate_duplicate_check/",
                          headers=self.headers, data=json.dumps(self.data, default=str), verify=False)

        time.sleep(1)
        resp_dict = json.loads(r.content)
        self.is_duplicate = resp_dict["IsDuplicate"]
        # print self.is_duplicate
        self.message = resp_dict['Message']

        if self.is_duplicate:
            self.is_duplicate1 = "Duplicate"
            # print self.is_duplicate1
        else:
            self.is_duplicate1 = "NotDuplicate"
            # print self.is_duplicate1

        if self.is_duplicate1 == self.current_data.get('ExpectedOutput'):
            self.style6 = self.style3
        else:
            self.style6 = self.style2

        self.excelWrite(self.message)

    def excelWrite(self, message):

        if message == self.current_data.get('Message'):
            self.status = "Pass"
            style1 = self.style3
        else:
            self.status = "Fail"
            style1 = self.style2

        self.ws.write(self.rowsize, 0, self.current_data.get('CandidateName'), self.style1)
        self.ws.write(self.rowsize, 1, self.current_data.get('FirstName'), self.style1)
        self.ws.write(self.rowsize, 2, self.current_data.get('MiddleName'), self.style1)
        self.ws.write(self.rowsize, 3, self.current_data.get('LastName'), self.style1)
        self.ws.write(self.rowsize, 4, self.current_data.get('EmailAddress'), self.style1)
        self.ws.write(self.rowsize, 5, self.current_data.get('MobileNumber'), self.style1)
        self.ws.write(self.rowsize, 6, self.current_data.get('PhoneNumber'), self.style1)
        self.ws.write(self.rowsize, 7, self.current_data.get('MaritalStatus'), self.style1)
        self.ws.write(self.rowsize, 8, self.current_data.get('Gender'), self.style1)
        self.ws.write(self.rowsize, 9, self.date_of_birth, self.style1)
        self.ws.write(self.rowsize, 10, self.current_data.get('Pancard'), self.style1)
        self.ws.write(self.rowsize, 11, self.current_data.get('Passport'), self.style1)
        self.ws.write(self.rowsize, 12, self.current_data.get('Aadhar'), self.style1)
        self.ws.write(self.rowsize, 13, self.current_data.get('USN'), self.style1)
        self.ws.write(self.rowsize, 14, self.current_data.get('College'), self.style1)
        self.ws.write(self.rowsize, 15, self.current_data.get('Degree'), self.style1)
        self.ws.write(self.rowsize, 16, self.current_data.get('Location'), self.style1)
        self.ws.write(self.rowsize, 17, self.current_data.get('TotalExperienceInMonths'), self.style1)

        self.ws.write(self.rowsize, 18, self.current_data.get('LinkedIn'), self.style1)
        self.ws.write(self.rowsize, 19, self.current_data.get('Facebook'), self.style1)
        self.ws.write(self.rowsize, 20, self.current_data.get('Twitter'), self.style1)

        self.ws.write(self.rowsize, 21, self.current_data.get('Text1'), self.style1)
        self.ws.write(self.rowsize, 22, self.current_data.get('Text2'), self.style1)
        self.ws.write(self.rowsize, 23, self.current_data.get('Text3'), self.style1)
        self.ws.write(self.rowsize, 24, self.current_data.get('Text4'), self.style1)
        self.ws.write(self.rowsize, 25, self.current_data.get('Text5'), self.style1)
        self.ws.write(self.rowsize, 26, self.current_data.get('Text6'), self.style1)
        self.ws.write(self.rowsize, 27, self.current_data.get('Text7'), self.style1)
        self.ws.write(self.rowsize, 28, self.current_data.get('Text8'), self.style1)
        self.ws.write(self.rowsize, 29, self.current_data.get('Text9'), self.style1)
        self.ws.write(self.rowsize, 30, self.current_data.get('Text10'), self.style1)
        self.ws.write(self.rowsize, 31, self.current_data.get('Text11'), self.style1)
        self.ws.write(self.rowsize, 32, self.current_data.get('Text12'), self.style1)
        self.ws.write(self.rowsize, 33, self.current_data.get('Text13'), self.style1)
        self.ws.write(self.rowsize, 34, self.current_data.get('Text14'), self.style1)
        self.ws.write(self.rowsize, 35, self.current_data.get('Text15'), self.style1)

        self.ws.write(self.rowsize, 36, self.current_data.get('Integer1'), self.style1)
        self.ws.write(self.rowsize, 37, self.current_data.get('Integer2'), self.style1)
        self.ws.write(self.rowsize, 38, self.current_data.get('Integer3'), self.style1)
        self.ws.write(self.rowsize, 39, self.current_data.get('Integer4'), self.style1)
        self.ws.write(self.rowsize, 40, self.current_data.get('Integer5'), self.style1)
        self.ws.write(self.rowsize, 41, self.current_data.get('Integer6'), self.style1)
        self.ws.write(self.rowsize, 42, self.current_data.get('Integer7'), self.style1)
        self.ws.write(self.rowsize, 43, self.current_data.get('Integer8'), self.style1)
        self.ws.write(self.rowsize, 44, self.current_data.get('Integer9'), self.style1)
        self.ws.write(self.rowsize, 45, self.current_data.get('Integer10'), self.style1)
        self.ws.write(self.rowsize, 46, self.current_data.get('Integer11'), self.style1)
        self.ws.write(self.rowsize, 47, self.current_data.get('Integer12'), self.style1)
        self.ws.write(self.rowsize, 48, self.current_data.get('Integer13'), self.style1)
        self.ws.write(self.rowsize, 49, self.current_data.get('Integer14'), self.style1)
        self.ws.write(self.rowsize, 50, self.current_data.get('Integer15'), self.style1)
        self.ws.write(self.rowsize, 51, self.current_data.get('DuplicationRuleText'), self.style1)

        self.ws.write(self.rowsize, 52, self.current_data.get('ExpectedOutput'), self.style6)
        self.ws.write(self.rowsize, 53, self.is_duplicate1, self.style6)
        self.ws.write(self.rowsize, 54, self.current_data.get('Message'), style1)
        self.ws.write(self.rowsize, 55, message, style1)
        self.ws.write(self.rowsize, 56, self.status, style1)

        self.rowsize = self.rowsize + 1
        self.wb_Result.save(
            '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/CRPO/DuplicationRule/'
            'DuplicationRule(' + self.__current_DateTime + ').xls')


ob = VerifyDuplicationRule()
