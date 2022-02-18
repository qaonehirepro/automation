import datetime
import os

path = os.getcwd()
# print(path)
# SCRIPTS\\API_SCRIPTS is current working directory so to use the previous directory below operation is done
# input_common_dir = path.replace('\SCRIPTS\COMMON', '\\PythonWorkingScripts_InputData\\')
# output_common_dir = path.replace('\SCRIPTS\COMMON', '\\PythonWorkingScripts_Output\\')
# chrome_driver_path = path.replace('\SCRIPTS\COMMON', '\\chromedriver.exe')
# print(input_common_dir)
# print(output_common_dir)
# print(chrome_driver_path)
input_common_dir = 'F:\\qa_automation\\PythonWorkingScripts_InputData\\'
output_common_dir = 'F:\\qa_automation\\PythonWorkingScripts_Output\\'
chrome_driver_path = 'F:\\qa_automation\\chromedriver.exe'
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

input_path_microsite_create_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'
input_path_microsite_update_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'
input_path_microsite_generic_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'

#security
input_path_ssrf_check = input_common_dir + 'SSRF\\SSRF_Final1.xls'
input_path_encryption_check = input_common_dir + 'Security\encryption.xls'

# UI Automation Input Path
input_path_ui_mcq_randomization = input_common_dir + 'UI\\Assessment\\qprandomization_automation.xls'
input_path_ui_assessment_verification = input_common_dir + 'UI\\Assessment\\ui_relogin.xls'
input_path_ui_qp_verification = input_common_dir + "UI\\Assessment\\qp_verification.xls"
input_path_ui_hirepro_chaining= input_common_dir + 'UI\\Assessment\\hirepro_chaining_at.xls'

# output paths
output_path_allowed_extension = output_common_dir + 'allowed_extensions\\API_allowed_extensions(' + started + ').xlsx'
output_path_applicant_report = output_common_dir + 'Assessment\\report\\API_applicantreport'
output_path_2tests_chaining = output_common_dir + 'Assessment\\API_2tests_Chaining_Automation -'
output_path_3tests_chaining = output_common_dir + 'Assessment\\API_3tests_Chaining_Automation - '
output_path_plagiarism_report = output_common_dir + 'Assessment\\plagiarism_report\\API_plagiarism_report'
output_path_proctor_evaluation = output_common_dir + 'Assessment\\proctoring\\API_proctoring_eval'
output_path_question_search_count = output_common_dir + 'Assessment\\search\\API_question_search_'
output_path_question_search_boundary = output_common_dir + 'Assessment\\search\\API_question_boundary_search_'
output_path_reinitiate_automation = output_common_dir + 'Assessment\\reinitiate\\API_reinitiate - '
output_path_ssrf_check = output_common_dir + 'SSRF\\API_security_check -'
output_path_microsite_create_case = output_common_dir + 'Microsite\\UI_Microsite_CreateCase(' + started + ').xlsx'
output_path_microsite_update_case = output_common_dir + 'Microsite\\UI_Microsite_UpdateCase(' + started + ').xlsx'
output_path_microsite_generic_case = output_common_dir + 'Microsite\\UI_Functionality_VandV(' + started + ').xlsx'
output_path_encryption_check = output_common_dir + 'SSRF\\API_encryption_check -'

# UI Automation Output Path
output_path_ui_vet_qs = output_common_dir + 'Assessment\\UI\\VET\\UI_qs(' + started + ').xlsx'
output_path_ui_cocubes = output_common_dir + 'UI\\UI_cocubes'
output_path_ui_mettl = output_common_dir + 'UI\\UI_Mettl'
output_path_ui_mcq_randomization = output_common_dir + 'UI\\UI_mcq_qprandomization_'
output_path_ui_subjective_randomization = output_common_dir + 'UI\\UI_subjective_qprandomization_'
output_path_ui_assessment_verification = output_common_dir + 'UI\\UI_ui_assessment_relogin.xls'
output_path_ui_qp_verification = output_common_dir + "UI\\UI_QP_verification"
output_path_ui_hirepro_chaining = output_common_dir + 'UI\\UI_hirepro_chaining - ' + started + '.xls'
# Assessment URLS

amsin_at_assessment_url = 'https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0='
amsin_automation_assessment_url = 'https://amsin.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF1dG9tYXRpb24ifQ=='
# print(input_path_ui_assessment_verification)