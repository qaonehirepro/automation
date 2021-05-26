import requests
import json
import mysql
import xlwt
import datetime
from all_configurations.api_lists import *
from mysql import connector

class Configurations():


    def __init__(self):
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")

    def amsDBConnectivity(self):

        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()
        return self.cursor


    @staticmethod
    def login():
        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "Automation", "UserName": "admin"}
        login_request = requests.post(api_object.login_api, headers=header,
                                 data=json.dumps(data), verify=True)
        login_response = login_request.json()
        login_token =  login_response['Token']
        return login_token

    def input_file_path(self):
        self.candidate_search_input_path = '/home/muttumurgan/Desktop/AutomationScripts/PythonWorkingScripts_InputData/' \
                                'CRPO/Candidate/Candidate_Combined_Search_Boundary_Condition.xls'


    def output_file_path(self):
        self.candidate_search_count_output_path = '/home/muttumurgan/Desktop/AutomationScripts/' \
                                            'PythonWorkingScripts_Output/CRPO/CandidateSearch/' \
                                            '%s_Candidate_Search_Only_Count.xls' %self.__current_DateTime


all_config_object = Configurations()
all_config_object.amsDBConnectivity()
all_config_object.input_file_path()
all_config_object.output_file_path()
