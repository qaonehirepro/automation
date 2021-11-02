from automation.Excel_Manipulation.COMMON.report import *
from automation.Excel_Manipulation.COMMON.writeExcel import *
from automation.Excel_Manipulation.CRPO.crpo_common import *
from automation.Excel_Manipulation.CRPO.credentials import *
from automation.Excel_Manipulation.COMMON.io_path import *
from automation.Excel_Manipulation.COMMON.api_requests_for_reports import *

crpo_headers = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                             cred_crpo_admin.get('tenant'))
for report_items in range(0, 1):
    try:
        report_obj.writeExcelConfigurations(output_path_plagiarism_report)
        download_api_response = crpo_common_obj.generate_plagiarism_report(crpo_headers, getall_plagarism_request_payload)
        report_obj.downloadReport(crpo_headers, input_path_plagiarism_report_downloaded, download_api_response)
        write_excel_object.save_result(output_path_plagiarism_report)
        write_excel_object.excelReadExpectedSheet(input_path_plagiarism_report)
        write_excel_object.excelReadActualSheet(input_path_plagiarism_report_downloaded)
        write_excel_object.excelWriteHeaders(hierarchy_headers_count=1)
        write_excel_object.excelMatchValues(usecase_name='Plagiarism Report', comparision_required_from_index=1,
                                            total_testcase_count=40)
    except Exception as e:
        print("Please Verify it manually...")
        print(e)
