from SCRIPTS.COMMON.write_excel_new import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.CRPO_COMMON.credentials import *
from SCRIPTS.COMMON.io_path import *


class AllowedFileExtensions:

    def __init__(self):
        self.allowed_extensions = [".doc", ".rtf", ".dot", ".docx", ".docm", ".dotx", ".dotm", ".docb", ".pdf", ".xls",
                                   ".xlt", ".xlm", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xla", ".xlam", ".xll",
                                   ".xlw", ".ppt", ".pot", ".pps", ".pptx", ".pptm", ".potx", ".potm", ".ppam", ".ppsx",
                                   ".ppsm", ".sldx", ".sldm", ".zip", ".rar", ".7z", ".gz", ".jpeg", ".jpg", ".gif",
                                   ".png", ".msg", ".txt", ".mp4", ".mvw", ".3gp", ".sql", ".webm", ".csv", ".odt",
                                   ".json", ".ods", ".ogg", ".p12"]
        requests.packages.urllib3.disable_warnings()
        self.row_size = 2
        write_excel_object.save_result(output_path_allowed_extension)
        header = ["Allowed Extensions"]
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        header1 = ["Extension Type", "Status", "File Name", "Expected Status", "Actual Status", "API Response"]
        write_excel_object.write_headers_for_scripts(1, 0, header1, write_excel_object.black_color_bold)

    def validate_files(self, token, excel_input):
        print(input_path_allowed_extension_files)
        print(excel_input.get('filePathName'))
        file_path = input_path_allowed_extension_files % (excel_input.get('filePathName'))
        file_name = excel_input.get('fileName')
        print(file_name)
        resp = crpo_common_obj.upload_files(token, file_name, file_path)
        if resp.get('status') == 'OK' and resp['data']:
            actual_status = 'allowed'
        elif "DisAllowed fileExtension" in (resp['error']['errorDescription']):
            actual_status = 'disallowed - fileExtension for the control'
        elif 'File extension does not match file format' in resp['error']['errorDescription']:
            actual_status = "disallowed - File extension does not match file format"
        elif 'UnAcceptable file format' in resp['error']['errorDescription']:
            actual_status = "UnAcceptable file format"
        elif 'DisAllowedFileExtFiles' in resp['error']['errorDescription']:
            actual_status = resp['error']['errorDescription']
        elif 'ExtNFileFormatMismatchedFiles' in resp['error']['errorDescription']:
            actual_status = resp['error']['errorDescription']
        else:
            actual_status = "status unknown"
        write_excel_object.compare_results_and_write_vertically(excel_input.get('extensionType'), None, self.row_size, 0)
        write_excel_object.compare_results_and_write_vertically(excel_input.get('fileName'), None, self.row_size, 2)
        write_excel_object.compare_results_and_write_vertically(excel_input.get('expectedStatus'), actual_status,
                                                                self.row_size, 3)
        write_excel_object.compare_results_and_write_vertically(write_excel_object.current_status, None, self.row_size, 1)
        self.row_size += 1


allowed_ext_obj = AllowedFileExtensions()
login_token = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                            cred_crpo_admin.get('tenant'))

excel_read_obj.excel_read(input_path_allowed_extension, 0)
excel_data = excel_read_obj.details
for data in excel_data:
    allowed_ext_obj.validate_files(login_token, data)
write_excel_object.write_overall_status(testcases_count=40)
