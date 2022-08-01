import mysql
import mysql.connector
import datetime


class delete_ssrf_data:

    def __init__(self):
        print(datetime.datetime.now())

    def amsdbconnection(self):
        # replica = 35.154.213.175
        # master = 35.154.36.218
        self.conn = mysql.connector.connect(host='35.154.213.175',
                                            database='appserver_core',
                                            user='qauser',
                                            password='qauser')
        self.cursor = self.conn.cursor()

    def commit_changes(self, query):
        pass

    def delete_assessment_test_users(self):
        self.amsdbconnection()
        delete_candidates = 'delete from candidates where id in (select candidate_id ' \
                            'from test_users where test_id  in (14671,14673,14675,14677));'
        print(delete_candidates)

        self.cursor.execute(delete_candidates)
        self.conn.commit()
        delete_test_users = 'delete from test_users where test_id  in (14671,14673,14675,14677);'
        print(delete_test_users)
        self.cursor.execute(delete_test_users)
        self.conn.commit()
        self.conn.close()


del_data = delete_ssrf_data()
del_data.delete_assessment_test_users()
print(datetime.datetime.now())
