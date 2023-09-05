from SCRIPTS.CRPO_COMMON.credentials import *
from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.write_excel_new import *
from SCRIPTS.COMMON.io_path import *


class IsStrictValidations:

    def __init__(self):
        self.row_count = 2
        write_excel_object.save_result(output_path_adhoc_profile_api_validation)
        # 0th Row Header
        header = ['Adhoc Profile API Security Validation']
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        # 1st Row Header
        header = ['TestCases', 'Status', 'Tenant Name', 'typeOfUser', 'user name', 'Is Adhoc enabled',
                  'Exp – Allowed Api Call status', 'Act – Allowed API Call status', 'Exp – Not allowed Api Call status',
                  'Act – Not Allowed API Call status']
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

        # No_profile = {"Role": {"RoleId": 15755, "RoleName": "MS_Automation_Role", "Description": "MS Automation Role",
        #                        "IsSystemRole": False, "AdhocProfileId": None, "DefaultExpiryInDays": None}}
        #
        # With_Profile = {"Role": {"RoleId": 15755, "RoleName": "MS_Automation_Role", "Description": "MS Automation Role",
        #                          "IsSystemRole": False, "AdhocProfileId": 43, "DefaultExpiryInDays": None}}

    def domain_validations(self, admin_token,  xl_data):
        write_excel_object.current_status_color = write_excel_object.green_color
        write_excel_object.current_status = "Pass"
        if xl_data.get('IsAdhocEnabled') == 'No':
            request = {"Role": {"RoleId": 15755, "RoleName": "MS_Automation_Role", "Description": "MS Automation Role",
                                "IsSystemRole": False, "AdhocProfileId": None, "DefaultExpiryInDays": None}}
            update_role = CrpoCommon.update_role(request, admin_token)
            print(update_role)
            adhoc_token = crpo_common_obj.login_to_crpo(amsin_at_adhoc_profile.get('user'),
                                                           amsin_at_adhoc_profile.get('password'),
                                                           amsin_at_adhoc_profile.get('tenant'))
            allowed_api = CrpoCommon.auth_user_v2(adhoc_token)
            allowed_api_status = allowed_api.get('status')
            not_allowed_api = CrpoCommon.get_app_preference(adhoc_token)
            print(not_allowed_api)
            not_allowed_api_status = not_allowed_api.get('status')
            # print(not_allowed_api)
            if not_allowed_api_status is None:
                not_allowed_api_status = 'OK'
            elif not_allowed_api_status == 'KO':
                not_allowed_api_status = not_allowed_api.get('error').get('errorMessage')
            # print(allowed_api_status)
            # print(not_allowed_api_status)

        else:
            print("This is else part")
            request = {"Role": {"RoleId": 15755, "RoleName": "MS_Automation_Role", "Description": "MS Automation Role",
                                "IsSystemRole": False, "AdhocProfileId": 43, "DefaultExpiryInDays": None}}
            update_role = CrpoCommon.update_role(request, admin_token)
            # print(update_role)
            adhoc_token = crpo_common_obj.login_to_crpo(amsin_at_adhoc_profile.get('user'),
                                                        amsin_at_adhoc_profile.get('password'),
                                                        amsin_at_adhoc_profile.get('tenant'))
            allowed_api = CrpoCommon.auth_user_v2(adhoc_token)
            allowed_api_status = allowed_api.get('status')
            not_allowed_api = CrpoCommon.get_app_preference(adhoc_token)
            print (not_allowed_api)
            not_allowed_api_status = not_allowed_api.get('status')
            if not_allowed_api_status == 'KO':
                not_allowed_api_status = not_allowed_api.get('error').get('errorMessage')
            # print(allowed_api_status)
            # print(not_allowed_api_status)

        write_excel_object.compare_results_and_write_vertically(xl_data.get('expConfiguredAPIStatus'),
                                                                allowed_api_status, self.row_count, 6)
        print(not_allowed_api_status)
        write_excel_object.compare_results_and_write_vertically(xl_data.get('expNotConfiguredAPIStatus'),
                                                                not_allowed_api_status, self.row_count, 8)
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('testCases'), None,
                                                                self.row_count, 0)
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('tenantName'), None,
                                                                self.row_count, 2)
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('typeOfUser'), None,
                                                                self.row_count, 3)
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('user'), None, self.row_count, 4)
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('IsAdhocEnabled'), None,
                                                                self.row_count, 5)
        write_excel_object.compare_results_and_write_vertically(write_excel_object.current_status, None, self.row_count,
                                                                1)
        self.row_count = self.row_count + 1


validation_obj = IsStrictValidations()
at_admin_token = crpo_common_obj.login_to_crpo(cred_crpo_admin_at.get('user'), cred_crpo_admin_at.get('password'),
                                               cred_crpo_admin_at.get('tenant'))
# print(at_admin_token)

excel_read_obj.excel_read(input_adhoc_profile_validations, 0)
excel_data = excel_read_obj.details
for current_excel_data in excel_data:
    validation_obj.domain_validations(at_admin_token, current_excel_data)
write_excel_object.write_overall_status(testcases_count=2)
