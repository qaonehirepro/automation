# This  script is for hirepro cron data clean up
import mysql
import mysql.connector
import datetime


class DeleteQuestionStatistics:

    def __init__(self):
        print(datetime.datetime.now())

    def amsdbconnection(self):
        self.conn = mysql.connector.connect(host='35.154.213.175',
                                            database='appserver_core',
                                            user='qauser',
                                            password='qauser')
        self.cursor = self.conn.cursor()

    def commit_changes(self, query):
        pass

    def delete_question_statistics(self):
        self.amsdbconnection()
        question_statistics = "delete from question_statisticss where id in " \
                              "(select questionstatistics_id from questions " \
                              "where id in (132829,132837,132839,132833,132835,132843,132845,132841,132847,132849,132851,132853,132861,132863,132857,132859,132865,132867,132869,132871,132873) and tenant_id in (1,248, 1787));"
        self.cursor.execute(question_statistics)
        self.conn.commit()

        tuser_results = "update test_results set is_statistics_completed = 0 " \
                        "where question_id in (132829,132837,132839,132833,132835,132843,132845,132841,132847,132849,132851,132853,132861,132863,132857,132859,132865,132867,132869,132871,132873) and question_tenant_id in (1,248, 1787);"
        self.cursor.execute(tuser_results)
        self.conn.commit()

        questions = "update questions set questionstatistics_id=NULL " \
                    "where id in (132829,132837,132839,132833,132835,132843,132845,132841,132847,132849,132851,132853,132861,132863,132857,132859,132865,132867,132869,132871,132873) and tenant_id in (1,248, 1787);"

        self.cursor.execute(questions)
        self.conn.commit()
        self.conn.close()


del_data = DeleteQuestionStatistics()
del_data.delete_question_statistics()
print(datetime.datetime.now())
