from SCRIPTS.COMMON.report import *
from SCRIPTS.COMMON.write_excel_new import *
from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.CRPO_COMMON.credentials import *
from SCRIPTS.COMMON.io_path import *
from SCRIPTS.COMMON.api_requests_for_reports import *

crpo_headers = crpo_common_obj.login_to_crpo(cred_crpo_admin.get('user'), cred_crpo_admin.get('password'),
                                             cred_crpo_admin.get('tenant'))
for report_items in range(0, 1):
    try:
        report_obj.writeExcelConfigurations(output_path_applicant_report)
        download_api_response = crpo_common_obj.generate_applicant_report(crpo_headers, getall_applicant_request_payload)
        report_obj.downloadReport(crpo_headers, input_path_applicant_report_downloaded, download_api_response)
        write_excel_object.save_result(output_path_applicant_report)
        write_excel_object.excelReadExpectedSheet(input_path_applicant_report)
        write_excel_object.excelReadActualSheet(input_path_applicant_report_downloaded)
        write_excel_object.excelWriteHeaders(hierarchy_headers_count=3)
        write_excel_object.excelMatchValues(usecase_name='Applicant Report', comparision_required_from_index=1,
                                            total_testcase_count=48)
    except Exception as e:
        print("Please Verify it manually...")
        print(e)
