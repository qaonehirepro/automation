from SCRIPTS.COMMON.dbconnection import *
import datetime


class EncryptionDelete:

    def __init__(self):
        print(datetime.datetime.now())

    @staticmethod
    def encryption_delete():
        db_connection = ams_db_connection()
        cursor = db_connection.cursor()
        cursor.execute("delete from candidates where hp_dec(usn) ='ENCRYPTION_CHECK2' and tenant_id=248;")
        db_connection.commit()
        db_connection.close()


del_data = EncryptionDelete()
del_data.encryption_delete()
print(datetime.datetime.now())
