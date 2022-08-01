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

# Assessment URLS
common_domain = 'https://amsin'
amsin_at_assessment_url = common_domain + ".hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0="
amsin_at_vet_url = 'https://pearsonstg.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF0In0='
amsin_automation_assessment_url = common_domain + '.hirepro.in/assessment/#/assess/login/eyJhbGlhcyI6ImF1dG9tYXRpb24ifQ=='
# print(input_path_ui_assessment_verification)

input_common_dir = 'F:\\qa_automation\\PythonWorkingScripts_InputData\\'
output_common_dir = 'F:\\qa_automation\\PythonWorkingScripts_Output\\'
chrome_driver_path = 'F:\\qa_automation\\chromedriver.exe'
started = datetime.datetime.now()
started = started.strftime("%d-%m-%Y")
# input paths
input_path_allowed_extension = input_common_dir + 'allowed_extensions\\allowed_extensions_inputfile.xls'
input_path_allowed_extension_files = input_common_dir + 'allowed_extensions\\%s'
input_path_applicant_report = input_common_dir + 'Assessment\\applicant_report\\applicantreport.xlsx'
input_path_applicant_report_downloaded = input_common_dir + 'Assessment\\applicant_report\\downloaded\\downloadedfile.xlsx'
input_path_2tests_chaining = input_common_dir + 'Assessment\\chaining\\2ndlogincase.xls'
# input_path_2tests_chaining = input_common_dir + 'Assessment\\chaining\\2ndlogincase - Copy.xls'
input_path_3tests_chaining = input_common_dir + 'Assessment\\chaining\\3_tests_login_automation.xls'
# input_path_3tests_chaining = input_common_dir + 'Assessment\\chaining\\3_tests_login_automationFailed.xls'
input_path_plagiarism_report = input_common_dir + 'Assessment\\plagiarism_report\\plagiarism_report.xlsx'
input_path_plagiarism_report_downloaded = input_common_dir + 'Assessment\\plagiarism_report\\downloaded\\plagiarism_report' + started + '.xlsx'
input_path_proctor_evaluation = input_common_dir + 'Assessment\\proc_eval\\proc_eval3.xls'
input_path_question_search_count = input_common_dir + 'Assessment\\Search\\question_search_Automation.xls'
input_path_question_search_boundary = input_common_dir + 'Assessment\\Search\\question_search_boundary_automation.xls'
input_path_reinitiate_automation = input_common_dir + 'Assessment\\reinitiateautomation1.xls'
input_path_mic_distortion_check = input_common_dir + 'Assessment\\mic_distortion_check\\1input.xls'
input_path_mic_distortion_files = input_common_dir + 'Assessment\\mic_distortion_check\\%s'
input_path_brightness_check = input_common_dir + 'Assessment\\brightnesscheck\\brightnesscheck.xls'
input_path_brightness_check_files = input_common_dir + 'Assessment\\brightnesscheck\\%s'
input_coding_compiler = input_common_dir + 'Assessment\\coding\\codingcompiler.xls'

input_path_microsite_create_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'
input_path_microsite_update_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'
input_path_microsite_generic_case = input_common_dir + 'Microsite\\GenericExcelTest.xls'

# security
input_path_ssrf_check = input_common_dir + 'SSRF\\SSRF_Final1.xls'
input_path_encryption_check = input_common_dir + 'Security\encryption.xls'

# UI Automation Input Path
input_path_ui_mcq_randomization = input_common_dir + 'UI\\Assessment\\qprandomization_automation.xls'
input_path_ui_assessment_verification = input_common_dir + 'UI\\Assessment\\ui_relogin.xls'
input_path_ui_qp_verification = input_common_dir + "UI\\Assessment\\qp_verification.xls"
input_path_ui_hirepro_chaining = input_common_dir + 'UI\\Assessment\\hirepro_chaining_at.xls'
input_path_ui_vet_vet_chaining = input_common_dir + 'UI\\Assessment\\vet_chaining.xls'
input_path_ui_test_security = input_common_dir + 'UI\\Assessment\\test_security.xls'
input_path_ui_reuse_score = input_common_dir + 'Assessment\\reuse_score.xls'
input_path_ui_mcq_client_section_random = input_common_dir + 'UI\\Assessment\\clientside_randomization.xls'

# output paths
output_path_allowed_extension = output_common_dir + 'allowed_extensions\\API_allowed_extensions(' + started + ').xlsx'
output_path_applicant_report = output_common_dir + 'Assessment\\report\\API_applicantreport'
output_path_2tests_chaining = output_common_dir + 'Assessment\\API_2tests_Chaining_Automation -'
output_path_3tests_chaining = output_common_dir + 'Assessment\\API_3tests_Chaining_Automation - '
output_path_plagiarism_report = output_common_dir + 'Assessment\\plagiarism_report\\API_plagiarism_report'
output_path_proctor_evaluation = output_common_dir + 'Assessment\\proctoring\\API_proctoring_eval100'
output_path_question_search_count = output_common_dir + 'Assessment\\search\\API_question_search_'
output_path_question_search_boundary = output_common_dir + 'Assessment\\search\\API_question_boundary_search_'
output_path_reinitiate_automation = output_common_dir + 'Assessment\\reinitiate\\API_reinitiate - '
output_path_ssrf_check = output_common_dir + 'SSRF\\API_security_check -'
output_path_microsite_create_case = output_common_dir + 'Microsite\\UI_Microsite_CreateCase(' + started + ').xlsx'
output_path_microsite_update_case = output_common_dir + 'Microsite\\UI_Microsite_UpdateCase(' + started + ').xlsx'
output_path_microsite_generic_case = output_common_dir + 'Microsite\\UI_Functionality_VandV(' + started + ').xlsx'
output_path_encryption_check = output_common_dir + 'SSRF\\API_encryption_check -'
output_path_reuse_score = output_common_dir + 'Assessment\\reuse_score - '
output_path_mic_check = output_common_dir + 'Assessment\\mic_distortion_check'
output_path_brightness_sharpness_check = output_common_dir + 'Assessment\\API_brightness_sharpnesscheck- '
output_coding_compiler = output_common_dir + 'Assessment\\codingcompiler'

# UI Automation Output Path
output_path_ui_vet_qs = output_common_dir + 'Assessment\\UI\\VET\\UI_qs(' + started + ').xlsx'
output_path_ui_cocubes = output_common_dir + 'UI\\UI_cocubes'
output_path_ui_mettl = output_common_dir + 'UI\\UI_Mettl'
output_path_ui_wheebox = output_common_dir + 'UI\\UI_Wheebox'
output_path_ui_mcq_randomization = output_common_dir + 'UI\\UI_mcq_qprandomization_'
output_path_ui_coding_randomization = output_common_dir + 'UI\\UI_coding_qprandomization_'
output_path_ui_subjective_randomization = output_common_dir + 'UI\\UI_subjective_qprandomization_'
output_path_ui_assessment_verification = output_common_dir + 'UI\\UI_ui_assessment_relogin.xls'
output_path_ui_qp_verification = output_common_dir + "UI\\UI_QP_verification"
output_path_ui_test_security = output_common_dir + "UI\\UI_Test_Security"
output_path_ui_hirepro_chaining = output_common_dir + 'UI\\UI_hirepro_chaining - ' + started + '.xls'
output_path_ui_vet_vet_chaining = output_common_dir + 'UI\\UI_vet_vet_chaining_plus_retest_consent - ' + started + '.xls'
output_path_ui_mcq_client_section_random = output_common_dir + 'UI\\UI_client_mcq_random_sectionwise_qprandomization_'
output_path_ui_mcq_client_group_random = output_common_dir + 'UI\\UI_client_mcq_random_groupwise_qprandomization_'
output_path_ui_mcq_client_test_random = output_common_dir + 'UI\\UI_client_mcq_random_testlevel_qprandomization_'