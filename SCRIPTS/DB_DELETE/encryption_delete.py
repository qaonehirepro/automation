import mysql
import mysql.connector
import datetime


class delete_ssrf_data:

    def __init__(self):
        print(datetime.datetime.now())

    def amsdbconnection(self):
        # replica = 35.154.213.175
        # master = 35.154.36.218
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='qauser',
                                            password='qauser')
        self.cursor = self.conn.cursor()

    def commit_changes(self, query):
        pass

    def encryption_delete(self):
        self.amsdbconnection()
        candidate = 'delete from candidates where hp_dec(first_name) like "%Encryption%" and hp_dec(email1)="qaonehirepro@gmail.com" and tenant_id=248';
        print(candidate)

        self.cursor.execute(candidate)
        self.conn.commit()
        self.conn.close()


del_data = delete_ssrf_data()
del_data.encryption_delete()
print(datetime.datetime.now())
