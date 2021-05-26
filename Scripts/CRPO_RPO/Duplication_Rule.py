import requests
import json
from Duplication_Rule_Excel_Data import *
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

    def excelWriteHeader(self):

        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('EC_Verification')
        self.ws.write(0, 0, 'Name', self.__style0)
        self.ws.write(0, 1, 'Fname', self.__style0)
        self.ws.write(0, 2, 'Mname', self.__style0)
        self.ws.write(0, 3, 'Lname', self.__style0)
        self.ws.write(0, 4, 'Email Address', self.__style0)
        self.ws.write(0, 5, 'Mobile', self.__style0)
        self.ws.write(0, 6, 'Phone', self.__style0)
        self.ws.write(0, 7, 'Marital Status', self.__style0)
        self.ws.write(0, 8, 'Gender', self.__style0)
        self.ws.write(0, 9, 'DOB', self.__style0)
        self.ws.write(0, 10, 'PANCARD', self.__style0)
        self.ws.write(0, 11, 'PASSPORT', self.__style0)
        self.ws.write(0, 12, 'Aadhar', self.__style0)
        self.ws.write(0, 13, 'USN', self.__style0)
        self.ws.write(0, 14, 'College', self.__style0)
        self.ws.write(0, 15, 'Degree', self.__style0)
        self.ws.write(0, 16, 'Location', self.__style0)
        self.ws.write(0, 17, 'Total Experience(in Months)', self.__style0)

        self.ws.write(0, 18, 'LinkedIn', self.__style0)
        self.ws.write(0, 19, 'Facebook', self.__style0)
        self.ws.write(0, 20, 'Twitter', self.__style0)

        self.ws.write(0, 21, 'Text1', self.__style0)
        self.ws.write(0, 22, 'Text2', self.__style0)
        self.ws.write(0, 23, 'Text3', self.__style0)
        self.ws.write(0, 24, 'Text4', self.__style0)
        self.ws.write(0, 25, 'Text5', self.__style0)
        self.ws.write(0, 26, 'Text6', self.__style0)
        self.ws.write(0, 27, 'Text7', self.__style0)
        self.ws.write(0, 28, 'Text8', self.__style0)
        self.ws.write(0, 29, 'Text9', self.__style0)
        self.ws.write(0, 30, 'Text10', self.__style0)
        self.ws.write(0, 31, 'Text11', self.__style0)
        self.ws.write(0, 32, 'Text12', self.__style0)
        self.ws.write(0, 33, 'Text13', self.__style0)
        self.ws.write(0, 34, 'Text14', self.__style0)
        self.ws.write(0, 35, 'Text15', self.__style0)

        self.ws.write(0, 36, 'Integer1', self.__style0)
        self.ws.write(0, 37, 'Integer2', self.__style0)
        self.ws.write(0, 38, 'Integer3', self.__style0)
        self.ws.write(0, 39, 'Integer4', self.__style0)
        self.ws.write(0, 40, 'Integer5', self.__style0)
        self.ws.write(0, 41, 'Integer6', self.__style0)
        self.ws.write(0, 42, 'Integer7', self.__style0)
        self.ws.write(0, 43, 'Integer8', self.__style0)
        self.ws.write(0, 44, 'Integer9', self.__style0)
        self.ws.write(0, 45, 'Integer10', self.__style0)
        self.ws.write(0, 46, 'Integer11', self.__style0)
        self.ws.write(0, 47, 'Integer12', self.__style0)
        self.ws.write(0, 48, 'Integer13', self.__style0)
        self.ws.write(0, 49, 'Integer14', self.__style0)
        self.ws.write(0, 50, 'Integer15', self.__style0)
        self.ws.write(0, 51, 'Duplicate Rule', self.__style0)
        self.ws.write(0, 52, 'Expected Message', self.__style0)
        self.ws.write(0, 53, 'Actual Message', self.__style0)
        self.ws.write(0, 54, 'Status', self.__style0)
        # self.ws.write(0, 55, 'Message', self.__style0)

    def updateDuplicateRule(self, i):
        self.update_json_data = dup_ob.xl_duplication_rule_json[i]
        print self.update_json_data
        self.data1 = {"AppPreference": {"Id": 3595, "Content": self.update_json_data,
                                        "Type": "duplication_conf.default"}, "IsTenantGlobal": "true"}
        r = requests.post("https://amsin.hirepro.in/py/common/common_app_utils/save_app_preferences/",
                          headers=self.headers, data=json.dumps(self.data1, default=str), verify=False)
        # print r.content
        # print r.status_code

    def checkDuplicate(self, i):
        self.data = {"FirstName": dup_ob.xl_candidate_fname[i], "MiddleName": dup_ob.xl_candidate_mname[i],
                     "LastName": dup_ob.xl_candidate_lname[i],
                     "Email1": dup_ob.xl_email[i], "Mobile1": dup_ob.xl_mobile[i],
                     "PhoneOffice": dup_ob.xl_phone_office[i],
                     "MaritalStatus": dup_ob.xl_marital[i], "Gender": dup_ob.xl_gender[i],
                     "DateOfBirth": dup_ob.xl_dob[i],
                     "PanNo": dup_ob.xl_pancard[i], "PassportNo": dup_ob.xl_passport[i],
                     "AadhaarNo": dup_ob.xl_aadhar[i],
                     "CollegeId": dup_ob.xl_college[i], "DegreeId": dup_ob.xl_degree[i], "USN": dup_ob.xl_usn[i],
                     "CurrentLocationId": dup_ob.xl_current_Location[i],
                     "TotalExperience": dup_ob.xl_total_experience[i],
                     "FacebookLink": dup_ob.xl_facebook[i],
                     "TwitterLink": dup_ob.xl_twitter[i],
                     "LinkedInLink": dup_ob.xl_linkedin[i],
                     "Integer1": dup_ob.xl_integer1[i], "Integer2": dup_ob.xl_integer2[i],
                     "Integer3": dup_ob.xl_integer3[i],
                     "Integer4": dup_ob.xl_integer4[i], "Integer5": dup_ob.xl_integer5[i],
                     "Integer6": dup_ob.xl_integer6[i],
                     "Integer7": dup_ob.xl_integer7[i], "Integer8": dup_ob.xl_integer8[i],
                     "Integer9": dup_ob.xl_integer9[i],
                     "Integer10": dup_ob.xl_integer10[i], "Integer11": dup_ob.xl_integer11[i],
                     "Integer12": dup_ob.xl_integer12[i],
                     "Integer13": dup_ob.xl_integer13[i], "Integer14": dup_ob.xl_integer14[i],
                     "Integer15": dup_ob.xl_integer15[i],
                     "Text1": dup_ob.xl_text1[i], "Text2": dup_ob.xl_text2[i], "Text3": dup_ob.xl_text3[i],
                     "Text4": dup_ob.xl_text4[i], "Text5": dup_ob.xl_text5[i], "Text6": dup_ob.xl_text6[i],
                     "Text7": dup_ob.xl_text7[i], "Text8": dup_ob.xl_text8[i], "Text9": dup_ob.xl_text9[i],
                     "Text10": dup_ob.xl_text10[i], "Text11": dup_ob.xl_text11[i], "Text12": dup_ob.xl_text12[i],
                     "Text13": dup_ob.xl_text13[i], "Text14": dup_ob.xl_text14[i], "Text15": dup_ob.xl_text15[i]
                     }
        r = requests.post("https://amsin.hirepro.in/py/rpo/candidate_duplicate_check/",
                          headers=self.headers, data=json.dumps(self.data, default=str), verify=False)

        time.sleep(1)
        resp_dict = json.loads(r.content)
        self.is_duplicate = resp_dict["IsDuplicate"]
        self.message = resp_dict['Message']
        self.excelWrite(self.message, i)

    def excelWrite(self, message, i):

        if message == dup_ob.xl_expected_message[i]:
            self.status = "Pass"
            style1 = self.__style3
        else:
            self.status = "Fail"
            style1 = self.__style2

        self.ws.write(self.rowsize, 0, dup_ob.xl_candidate_name[i], self.__style1)
        self.ws.write(self.rowsize, 1, dup_ob.xl_candidate_fname[i], self.__style1)
        self.ws.write(self.rowsize, 2, dup_ob.xl_candidate_mname[i], self.__style1)
        self.ws.write(self.rowsize, 3, dup_ob.xl_candidate_lname[i], self.__style1)
        self.ws.write(self.rowsize, 4, dup_ob.xl_email[i], self.__style1)
        self.ws.write(self.rowsize, 5, dup_ob.xl_mobile[i], self.__style1)
        self.ws.write(self.rowsize, 6, dup_ob.xl_phone_office[i], self.__style1)
        self.ws.write(self.rowsize, 7, dup_ob.xl_marital[i], self.__style1)
        self.ws.write(self.rowsize, 8, dup_ob.xl_gender[i], self.__style1)
        self.ws.write(self.rowsize, 9, dup_ob.xl_dob[i], self.__style1)
        self.ws.write(self.rowsize, 10, dup_ob.xl_pancard[i], self.__style1)
        self.ws.write(self.rowsize, 11, dup_ob.xl_passport[i], self.__style1)
        self.ws.write(self.rowsize, 12, dup_ob.xl_aadhar[i], self.__style1)
        self.ws.write(self.rowsize, 13, dup_ob.xl_usn[i], self.__style1)
        self.ws.write(self.rowsize, 14, dup_ob.xl_college[i], self.__style1)
        self.ws.write(self.rowsize, 15, dup_ob.xl_degree[i], self.__style1)
        self.ws.write(self.rowsize, 16, dup_ob.xl_current_Location[i], self.__style1)
        self.ws.write(self.rowsize, 17, dup_ob.xl_total_experience[i], self.__style1)

        self.ws.write(self.rowsize, 18, dup_ob.xl_linkedin[i], self.__style1)
        self.ws.write(self.rowsize, 19, dup_ob.xl_facebook[i], self.__style1)
        self.ws.write(self.rowsize, 20, dup_ob.xl_twitter[i], self.__style1)

        self.ws.write(self.rowsize, 21, dup_ob.xl_text1[i], self.__style1)
        self.ws.write(self.rowsize, 22, dup_ob.xl_text2[i], self.__style1)
        self.ws.write(self.rowsize, 23, dup_ob.xl_text3[i], self.__style1)
        self.ws.write(self.rowsize, 24, dup_ob.xl_text4[i], self.__style1)
        self.ws.write(self.rowsize, 25, dup_ob.xl_text5[i], self.__style1)
        self.ws.write(self.rowsize, 26, dup_ob.xl_text6[i], self.__style1)
        self.ws.write(self.rowsize, 27, dup_ob.xl_text7[i], self.__style1)
        self.ws.write(self.rowsize, 28, dup_ob.xl_text8[i], self.__style1)
        self.ws.write(self.rowsize, 29, dup_ob.xl_text9[i], self.__style1)
        self.ws.write(self.rowsize, 30, dup_ob.xl_text10[i], self.__style1)
        self.ws.write(self.rowsize, 31, dup_ob.xl_text11[i], self.__style1)
        self.ws.write(self.rowsize, 32, dup_ob.xl_text12[i], self.__style1)
        self.ws.write(self.rowsize, 33, dup_ob.xl_text13[i], self.__style1)
        self.ws.write(self.rowsize, 34, dup_ob.xl_text14[i], self.__style1)
        self.ws.write(self.rowsize, 35, dup_ob.xl_text15[i], self.__style1)

        self.ws.write(self.rowsize, 36, dup_ob.xl_integer1[i], self.__style1)
        self.ws.write(self.rowsize, 37, dup_ob.xl_integer2[i], self.__style1)
        self.ws.write(self.rowsize, 38, dup_ob.xl_integer3[i], self.__style1)
        self.ws.write(self.rowsize, 39, dup_ob.xl_integer4[i], self.__style1)
        self.ws.write(self.rowsize, 40, dup_ob.xl_integer5[i], self.__style1)
        self.ws.write(self.rowsize, 41, dup_ob.xl_integer6[i], self.__style1)
        self.ws.write(self.rowsize, 42, dup_ob.xl_integer7[i], self.__style1)
        self.ws.write(self.rowsize, 43, dup_ob.xl_integer8[i], self.__style1)
        self.ws.write(self.rowsize, 44, dup_ob.xl_integer9[i], self.__style1)
        self.ws.write(self.rowsize, 45, dup_ob.xl_integer10[i], self.__style1)
        self.ws.write(self.rowsize, 46, dup_ob.xl_integer11[i], self.__style1)
        self.ws.write(self.rowsize, 47, dup_ob.xl_integer12[i], self.__style1)
        self.ws.write(self.rowsize, 48, dup_ob.xl_integer13[i], self.__style1)
        self.ws.write(self.rowsize, 49, dup_ob.xl_integer14[i], self.__style1)
        self.ws.write(self.rowsize, 50, dup_ob.xl_integer15[i], self.__style1)
        self.ws.write(self.rowsize, 51, dup_ob.xl_duplicatetext[i], self.__style1)
        self.ws.write(self.rowsize, 52, dup_ob.xl_expected_message[i], style1)
        self.ws.write(self.rowsize, 53, message, style1)
        self.ws.write(self.rowsize, 54, self.status, style1)

        self.rowsize = self.rowsize + 1
        ob.wb_Result.save(
            '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/CRPO/DuplicationRule/'
            'DuplicationRule(' + self.__current_DateTime + ').xls')


ob = VerifyDuplicationRule()
ob.excelWriteHeader()
tot = len(dup_ob.xl_candidate_name)
for iter in range(0, tot):
    ob.updateDuplicateRule(iter)
    ob.checkDuplicate(iter)
