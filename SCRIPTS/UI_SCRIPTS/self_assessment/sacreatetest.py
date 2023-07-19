from SCRIPTS.CRPO_COMMON.credentials import cred_crpo_admin
from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.UI_COMMON.assessment_ui_common_v2 import *
import time
from SCRIPTS.UI_SCRIPTS.assessment_data_verification import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.io_path import *

class SaTest:
    def __init__(self):
        print("hi")

    def tenant_login(self, tuname, tpass):
        print("inside tenant_login")
        tenant = 'AT'
        amsin_crpodemo_tenant_url = 'https://amsin.hirepro.in/crpo/#/login/at'
        self.browser = assess_ui_common_obj.initiate_browser2(amsin_crpodemo_tenant_url, chrome_driver_path)
        print("browser opened")
        login_details = crpo_common_obj.login_to_crpo(tuname, tpass, tenant)
        if login_details == 'SUCCESS':
            print("login success")
            time.sleep(5)
            self.browser.quit()
            time.sleep(5)
        else:
            print("login failed due to below reason")
            print(login_details)


print(datetime.datetime.now())
assessment_obj = SaTest()
# input_file_path = r"F:\qa_automation\PythonWorkingScripts_InputData\UI\Assessment\ui_relogin.xls"
tuser_name = 'admin'
tpassword = 'At@2023$$'
assessment_obj.tenant_login(tuser_name, tpassword)
print("main2")

print(datetime.datetime.now())
