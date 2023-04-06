from SCRIPTS.CRPO_COMMON.credentials import *
from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.write_excel_new import *
from SCRIPTS.COMMON.io_path import *


class IsStrictValidations:

    def __init__(self):
        self.row_count = 2
        write_excel_object.save_result(output_path_infra_strict_domain)
        # 0th Row Header
        header = ['Infra is strict domain changes']
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        # 1st Row Header
        header = ['TestCases', 'Status', 'Tenant Name', 'Type Of User', 'user name', 'domain name',
                  'App Name', 'Exp – Api Call status', 'Act – API Call status', 'Exp - Python server', 'Act - Python server']
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

    def domain_validations(self, non_eu_token, eu_non_strict_token, eu_is_strict_token, xl_data):
        write_excel_object.current_status_color = write_excel_object.green_color
        write_excel_object.current_status = "Pass"
        api_call_status = None
        app_node = None
        domain = xl_data.get('domainName')
        if xl_data.get('appName') != 'EMPTY':
            print("This is app name case")
            non_eu_token.update({'APP-NAME': xl_data.get('appName')})
            eu_non_strict_token.update({'APP-NAME': xl_data.get('appName')})
            eu_is_strict_token.update({'APP-NAME': xl_data.get('appName')})
        else:
            print("This is no AppName case")
            eu_is_strict_token.pop('APP-NAME')

        if xl_data.get('tenantName') == 'isstricteu1':
            app_node = CrpoCommon.app_node_by_random_api(domain, eu_is_strict_token)
            api_status = CrpoCommon.get_app_preference(domain, eu_is_strict_token)
            if api_status.get('status') == 'KO':
                api_call_status = api_status.get('error').get('errorDescription')
            else:
                api_call_status = api_status.get('status')
        elif xl_data.get('tenantName') == 'pearsontesteu':
            app_node = CrpoCommon.app_node_by_random_api(domain, eu_non_strict_token)
            api_status = CrpoCommon.get_app_preference(domain, eu_non_strict_token)
            if api_status.get('status') == 'KO':
                api_call_status = api_status.get('error').get('errorDescription')
            else:
                api_call_status = api_status.get('status')
        elif xl_data.get('tenantName') == 'at':
            app_node = CrpoCommon.app_node_by_random_api(domain, non_eu_token)
            api_status = CrpoCommon.get_app_preference(domain, non_eu_token)
            if api_status.get('status') == 'KO':
                api_call_status = api_status.get('error').get('errorDescription')
            else:
                api_call_status = api_status.get('status')
        else:
            print("Please check the tenant name")
        if app_node:
            app_nodes = app_node.split('@')
        write_excel_object.compare_results_and_write_vertically(xl_data.get('expectedApiStatus'), api_call_status,
                                                                self.row_count, 7)
        write_excel_object.compare_results_and_write_vertically(xl_data.get('pythonAppNode'), app_nodes[0],
                                                                self.row_count, 9)
        print(current_excel_data.get('testCases'))
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('testCases'), None,
                                                                self.row_count, 0)
        print(current_excel_data.get('tenantName'))
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('tenantName'), None,
                                                                self.row_count, 2)
        print(current_excel_data.get('typeOfUser'))
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('typeOfUser'), None,
                                                                self.row_count, 3)
        print(current_excel_data.get('user'))
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('user'), None, self.row_count, 4)
        print(current_excel_data.get('domainName'))
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('domainName'), None,
                                                                self.row_count, 5)
        print(current_excel_data.get('appName'))
        write_excel_object.compare_results_and_write_vertically(current_excel_data.get('appName'), None, self.row_count,
                                                                6)
        write_excel_object.compare_results_and_write_vertically(write_excel_object.current_status, None, self.row_count,
                                                                1)
        self.row_count = self.row_count + 1


validation_obj = IsStrictValidations()
at_token = crpo_common_obj.login_to_crpo(mumbai_non_eu_at.get('user'), mumbai_non_eu_at.get('password'),
                                         mumbai_non_eu_at.get('tenant'))
pearsontesteu_token = crpo_common_obj.login_to_crpo(mumbai_eu_non_strict_pearsontesteu.get('user'),
                                                    mumbai_eu_non_strict_pearsontesteu.get('password'),
                                                    mumbai_eu_non_strict_pearsontesteu.get('tenant'))
isstrict_token = crpo_common_obj.login_to_crpo(mumbai_eu_strict_isstrict1.get('user'),
                                               mumbai_eu_strict_isstrict1.get('password'),
                                               mumbai_eu_strict_isstrict1.get('tenant'))

excel_read_obj.excel_read(input_infra_strict_domain_validations, 0)
excel_data = excel_read_obj.details
for current_excel_data in excel_data:
    validation_obj.domain_validations(at_token, pearsontesteu_token, isstrict_token, current_excel_data)
write_excel_object.write_overall_status(testcases_count=12)
