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
                              "where id in (132101,132097,132107,132121,132123,132127,132135,132149,132157,132133," \
                              "132181,132391,132395,132397,132399,132401,132403,132407,132409,132411,132413,132405) " \
                              "and tenant_id=1787);"
        self.cursor.execute(question_statistics)
        self.conn.commit()

        tuser_results = "update test_results set is_statistics_completed = 0 " \
                        "where question_id in (132101,132097,132107,132121,132123,132127,132135,132149,132157,132133," \
                        "132181,132391,132395,132397,132399,132401,132403,132407,132409,132411,132413,132405) " \
                        "and question_tenant_id=1787;"
        self.cursor.execute(tuser_results)
        self.conn.commit()

        questions = "update questions set questionstatistics_id=NULL " \
                    "where id in (132101,132097,132107,132121,132123,132127,132135,132149,132157,132133,132181," \
                    "132391,132395,132397,132399,132401,132403,132407,132409,132411,132413,132405) and tenant_id=1787;"
        # print(tuser_scores)

        self.cursor.execute(questions)
        self.conn.commit()
        self.conn.close()


del_data = DeleteQuestionStatistics()
del_data.delete_question_statistics()
print(datetime.datetime.now())
