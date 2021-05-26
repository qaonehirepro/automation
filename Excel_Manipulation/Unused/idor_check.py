import xlsxwriter
from COMMON.read_excel import *
import requests
import json
import datetime


class idorCheck:

    def __init__(self):
        requests.packages.urllib3.disable_warnings()
        self.started = datetime.datetime.now()
        self.started = self.started.strftime("%Y-%M-%d-%H-%M-%S")
        input_file_path = '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_InputData/SSRF/IDOR2.xls'
        excel_read_obj.excel_read(input_file_path, 0)
        self.row_size = 2
        self.write_excel = xlsxwriter.Workbook(
            '/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_Output/SSRF/IDOR - ' + self.started + '.xls')
        self.ws = self.write_excel.add_worksheet()
        self.black_color = self.write_excel.add_format({'font_color': 'black', 'font_size': 9})
        self.red_color = self.write_excel.add_format({'font_color': 'red', 'font_size': 9})
        self.green_color = self.write_excel.add_format({'font_color': 'green', 'font_size': 9})
        self.black_color_bold = self.write_excel.add_format({'font_color': 'black', 'bold': True, 'font_size': 9})
        self.over_all_status_pass = self.write_excel.add_format({'font_color': 'green', 'bold': True, 'font_size': 9})
        self.over_all_status_failed = self.write_excel.add_format({'font_color': 'red', 'bold': True, 'font_size': 9})
        self.over_all_status_color = self.over_all_status_pass
        self.over_all_status = 'Pass'

        self.ws.write(0, 0, "IDOR - Security Check", self.black_color_bold)
        self.ws.write(1, 0, "Testcases", self.black_color)
        self.ws.write(1, 1, "Status", self.black_color)
        self.ws.write(1, 2, "API", self.black_color)
        self.ws.write(1, 3, "Request", self.black_color)
        self.ws.write(1, 4, "Response", self.black_color)
        self.ws.write(1, 5, "Expected Status", self.black_color)
        self.ws.write(1, 6, "Actual Status", self.black_color)
        self.ws.write(1, 6, "Token used", self.black_color)

    @staticmethod
    def login_to_crpo(tenant_name, user_name, password):
        header = {"content-type": "application/json"}
        data = {"LoginName": user_name, "Password": password, "TenantAlias": tenant_name, "UserName": user_name}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                 data=json.dumps(data), verify=False)
        login_response = response.json()
        headers = {"content-type": "application/json", "X-AUTH-TOKEN": login_response.get("Token")}

        return headers

    def security_check(self, api, request, api_module, expected_status, token1, token2):
        if api_module == 'candidate1':
            headers = token1
        else:
            headers = token2
        print headers
        api_request = requests.post(api, headers=headers, data=request, verify=False)
        resp_dict = json.loads(api_request.content)
        if resp_dict.get('status') == 'OK':
            actual_status = 'authorized'
        else:
            actual_status = "unauthorized"

        self.ws.write(self.row_size, 0, "Testcase Data", self.black_color)
        self.ws.write(self.row_size, 2, api, self.black_color)
        self.ws.write(self.row_size, 3, request, self.black_color)
        self.ws.write(self.row_size, 4, str(resp_dict), self.black_color)
        self.ws.write(self.row_size, 7, headers.get('X-AUTH-TOKEN'), self.black_color)
        if expected_status == actual_status:
            self.ws.write(self.row_size, 1, "Pass", self.green_color)
            self.ws.write(self.row_size, 5, expected_status, self.green_color)
            self.ws.write(self.row_size, 6, actual_status, self.green_color)

        else:
            self.ws.write(self.row_size, 1, "Fail", self.red_color)
            self.over_all_status = 'Fail'
            self.over_all_status_color = self.over_all_status_failed
            self.ws.write(self.row_size, 5, expected_status, self.red_color)
            self.ws.write(self.row_size, 6, actual_status, self.red_color)

        self.row_size += 1


security_obj = idorCheck()
security_data_excel = excel_read_obj.details
candidate1 = security_obj.login_to_crpo('automation', 'S1N1J1E1V101100001@gmail.com', '4LWS-067')
candidate2 = security_obj.login_to_crpo('automation', 'S1N1J1E1V10199001@gmail.com', '4LWS-067')
for iter in security_data_excel:
    security_obj.security_check(iter.get('api'), iter.get('payLoad'), iter.get('apiModule'), iter.get('expectedStatus'),
                                candidate1, candidate2)
ended = datetime.datetime.now()
ended = "Ended:- %s" % ended.strftime("%Y-%M-%d-%H-%M-%S")
security_obj.ws.write(0, 1, security_obj.over_all_status, security_obj.over_all_status_color)
security_obj.ws.write(0, 2, 'Started:-' + security_obj.started, security_obj.black_color_bold)
security_obj.ws.write(0, 3, ended, security_obj.black_color_bold)
security_obj.write_excel.close()
