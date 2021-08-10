import requests
import xlsxwriter
import datetime


class AllConfigurations():

    def __init__(self):
        self.total_jobs = {}
        requests.packages.urllib3.disable_warnings()
        self.total_jobs = {"job1": 28135}
        self.started = datetime.datetime.now()
        self.started = self.started.strftime("%Y-%M-%d")

    # -----------------------------------------------------------------------------------------------------------------#
    # 1."expected_excel_sheet_path" is our expected sheet file path
    # 2."download_path" is the actualsheet file path (which will be downloaded from application after running the script)
    # 3."save_path" is having output path (which will be generated post completion of the script)
    # -----------------------------------------------------------------------------------------------------------------#
    def filePath(self):
        self.expected_excel_sheet_path = r'D:\automation\PythonWorkingScripts_InputData\Assessment\plagiarism_report\plagiarism_report.xlsx'

        self.download_path = r'D:\automation\PythonWorkingScripts_InputData\Assessment\plagiarism_report\downloaded\plagiarism_report' + self.started + '.xlsx'
        self.save_path = r'D:\automation\PythonWorkingScripts_Output\Assessment\plagiarism_report\plagiarism_report.xls'

    # -----------------------------------------------------------------------------------------------------------------#
    # 1.  xlsxwriter.Workbook is used to create Excel Workbook in the specified path
    # 2. "add_worksheet" is used to create work sheet in the created workbook
    # -----------------------------------------------------------------------------------------------------------------#
    def writeExcelConfigurations(self):
        self.write_excel = xlsxwriter.Workbook(config_obj.save_path)
        self.ws = self.write_excel.add_worksheet()
        self.black_color = self.write_excel.add_format({'font_color': 'black', 'font_size': 9})
        self.red_color = self.write_excel.add_format({'font_color': 'red', 'font_size': 9})
        self.green_color = self.write_excel.add_format({'font_color': 'green', 'font_size': 9})
        self.black_color_bold = self.write_excel.add_format({'font_color': 'black', 'bold': True, 'font_size': 9})

    # ------------------------------------------------------------------------------------------------------------------#
    # 1. This method is having all api requests
    # 2. Here, all requests are used for download Assessment Interview report API
    # ------------------------------------------------------------------------------------------------------------------#
    def apiRequests(self):
        self.getall_plagarism_data = {"testId": 11966, "plagiarismMode": 1}


config_obj = AllConfigurations()
config_obj.apiRequests()
