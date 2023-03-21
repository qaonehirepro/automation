import mysql
import mysql.connector
from SCRIPTS.CRPO_COMMON.credentials import *


def ams_db_connection():
    conn = mysql.connector.connect(host=amsin_master.get('host'),
                                   database=amsin_master.get('database'),
                                   user=amsin_master.get('user'),
                                   password=amsin_master.get('password'))
    return conn

