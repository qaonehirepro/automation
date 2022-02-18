import mysql
import mysql.connector
from SCRIPTS.CRPO_COMMON.credentials import *


def ams_db_connection():
    conn = mysql.connector.connect(host=amsin_master.get('host'),
                                   database=amsin_master.get('database'),
                                   user=amsin_master.get('user'),
                                   password=amsin_master.get('password'))
    # cursor = conn.cursor()
    # query = 'select first_name, middle_name,last_name, candidate_name,email1,email2,mobile1,' \
    #         'phone_office,pan_card,passport,aadhaar_no,usn,phone_residence from candidates where id=1324791;'
    # cursor.execute(query)
    # data = cursor.fetchone()
    # print(data)
    #conn.close()
    return conn

