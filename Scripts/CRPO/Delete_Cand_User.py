import datetime
import mysql
import time
import xlrd
from mysql import connector

class DeleteQuery:
    def __init__(self):

        self.now = datetime.datetime.now()
        self.password = raw_input('DB password ::')

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_candidateId = []
        self.xl_userId = []
        self.xl_testuserId = []

    def dbconnection(self):

        self.connection = mysql.connector.connect(host='35.154.36.218',
                                                  database='appserver_core',
                                                  user='qauser',
                                                  password=self.password)
        self.cursor = self.connection.cursor()

    def candidate_excel_data(self):

        workbook = xlrd.open_workbook('//home/muthumurugan/Desktop/'
                                      'Automation/PythonWorkingScripts_Output/'
                                      'CRPO/Upload Candidate/API_UploadCandidates.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[2]:
                self.xl_candidateId.append(int(rows[2]))

        self.xl_candidateId1 = tuple(self.xl_candidateId)
        # self.candidate_query = "UPDATE appserver_core.candidates SET tenant_id=0, is_archived=1," \
        #                        "is_deleted=1, is_draft=0 WHERE id in{};".format(self.xl_candidateId1)
        # self.cursor.execute(self.candidate_query)
        # print self.candidate_query

        self.candidate_query1 = "DELETE FROM duplicate_candidates_infos where candidate_id in{};" \
            .format(self.xl_candidateId1)
        self.cursor.execute(self.candidate_query1)
        print self.candidate_query1

        self.candidate_query2 = "DELETE FROM appserver_core.candidates WHERE tenant_id=1787 and id in{};" \
            .format(self.xl_candidateId1)
        self.cursor.execute(self.candidate_query2)
        print self.candidate_query2

        self.candidate_query3 = "DELETE FROM test_users WHERE candidate_id in{};" .format(self.xl_candidateId1)
        self.cursor.execute(self.candidate_query3)
        print self.candidate_query3

    def user_excel_data(self):

        workbook = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/'
                                      'PythonWorkingScripts_Output/CRPO/'
                                      'Create USer/API_Create_User.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[2]:
                self.xl_userId.append(int(rows[2]))

        self.xl_userId1 = tuple(self.xl_userId)
        self.user_query = "UPDATE appserver_core.users SET tenant_id='0', is_archived='1'," \
                          " is_deleted='1' WHERE id in{}".format(self.xl_userId1)
        self.cursor.execute(self.user_query)
        print self.user_query

    def test_user_excel_data(self):
        workbook = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/'
                                      'PythonWorkingScripts_InputData/CRPO/ScoreSheet/UploadScores.xls')
        sheet1 = workbook.sheet_by_index(1)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[0]:
                self.xl_testuserId.append(int(rows[0]))

        self.xl_testuserId1 = tuple(self.xl_testuserId)
        self.testuser_query = "delete from candidate_scores where testuser_id in{};" \
            .format(self.xl_testuserId1)
        self.cursor.execute(self.testuser_query)
        print self.testuser_query

        self.testuser_query1 = "UPDATE test_users SET total_score ='', percentage ='', status='' WHERE id in {};" \
            .format(self.xl_testuserId1)
        self.cursor.execute(self.testuser_query1)
        print self.testuser_query1


Object = DeleteQuery()
Object.dbconnection()
Object.candidate_excel_data()
Object.user_excel_data()
Object.test_user_excel_data()
Object.connection.commit()
Object.connection.close()
