import datetime

chrome_driver_path = r"F:\qa_automation\automation\chromedriver.exe"
input_common_dir = 'F:\\qa_automation\\automation\\PythonWorkingScripts_InputData\\'
output_common_dir = 'F:\\qa_automation\\automation\\PythonWorkingScripts_Output\\'
started = datetime.datetime.now()
started = started.strftime("%d-%m-%Y")
# input paths
input_path_allowed_extension = input_common_dir + 'allowed_extensions\\allowed_extensions_inputfile.xls'
input_path_allowed_extension_files = input_common_dir + 'allowed_extensions\\%s'
input_path_applicant_report = input_common_dir + 'Assessment\\applicant_report\\job1_1.xlsx'
input_path_applicant_report_downloaded = input_common_dir + 'Assessment\\applicant_report\\downloaded\\downloadedfile.xlsx'
input_path_2tests_chaining = input_common_dir + 'Assessment\\chaining\\2ndlogincase.xls'
# input_path_2tests_chaining = input_common_dir + 'Assessment\\chaining\\2ndlogincase - Copy.xls'
input_path_3tests_chaining = input_common_dir + 'Assessment\\chaining\\3_tests_login_automation.xls'
input_path_plagiarism_report = input_common_dir + 'Assessment\\plagiarism_report\\plagiarism_report.xlsx'
input_path_plagiarism_report_downloaded = input_common_dir + 'Assessment\\plagiarism_report\\downloaded\\plagiarism_report' + started + '.xlsx'
input_path_proctor_evaluation = input_common_dir + 'Assessment\\proc_eval\\proc_eval3.xls'
input_path_question_search_count = input_common_dir + 'Assessment\\Search\\question_search_Automation.xls'
input_path_question_search_boundary = input_common_dir + 'Assessment\\Search\\question_search_boundary_automation.xls'
input_path_reinitiate_automation = input_common_dir + 'Assessment\\reinitiateautomation1.xls'
input_path_ssrf_check = input_common_dir + 'SSRF\\SSRF_Final1.xls'
input_path_microsite_create_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'
input_path_microsite_update_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'
input_path_microsite_generic_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'

# UI Automation Input Path
input_path_ui_mcq_randomization = input_common_dir + 'UI\\Assessment\\qprandomization_automation.xls'
# output paths
output_path_allowed_extension = output_common_dir + 'allowed_extensions\\allowed_extensions(' + started + ').xlsx'
output_path_applicant_report = output_common_dir + 'Assessment\\report\\applicantreport'
output_path_2tests_chaining = output_common_dir + 'Assessment\\Chaining_Automation -'
output_path_3tests_chaining = output_common_dir + 'Assessment\\3tests_Chaining_Automation - '
output_path_plagiarism_report = output_common_dir + 'Assessment\\plagiarism_report\\plagiarism_report'
output_path_proctor_evaluation = output_common_dir + 'Assessment\\proctoring\\proctoring_eval'
output_path_question_search_count = output_common_dir + 'Assessment\\search\\question_search_'
output_path_question_search_boundary = output_common_dir + 'Assessment\\search\\question_boundary_search_'
output_path_reinitiate_automation = output_common_dir + 'Assessment\\reinitiate\\reinitiate - '
output_path_ssrf_check = output_common_dir + 'SSRF\\security_check -'
output_path_microsite_create_case = output_common_dir + 'Microsite\\Microsite_CreateCase(' + started + ').xlsx'
output_path_microsite_update_case = output_common_dir + 'Microsite\\Microsite_UpdateCase(' + started + ').xlsx'
output_path_microsite_generic_case = output_common_dir + 'Microsite\\UI_Functionality_VandV(' + started + ').xlsx'

# UI Automation Output Path
output_path_ui_vet_qs = output_common_dir + 'Assessment\\UI\\VET\\qs(' + started + ').xlsx'
output_path_ui_mcq_randomization = output_common_dir + 'UI\\qprandomization_'

# Assessment URLS

amsin_at_assessment_url = 'https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0='
amsin_automation_assessment_url = 'https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF1dG9tYXRpb24ifQ=='
