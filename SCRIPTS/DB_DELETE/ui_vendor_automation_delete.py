from SCRIPTS.COMMON.dbconnection import *
import datetime


class UIVendorAutomationDelete:

    def __init__(self):
        print(datetime.datetime.now())

    @staticmethod
    def delete_assessment_test_users():
        db_connection = ams_db_connection()
        cursor = db_connection.cursor()
        delete_candidates = 'delete from candidates where id in (select candidate_id ' \
                            'from test_users where test_id  in (14671,14673,14675,14677));'
        print(delete_candidates)
        cursor.execute(delete_candidates)
        db_connection.commit()
        delete_test_users = 'delete from test_users where test_id  in (14671,14673,14675,14677);'
        print(delete_test_users)
        cursor.execute(delete_test_users)
        db_connection.commit()
        db_connection.close()


del_data = UIVendorAutomationDelete()
del_data.delete_assessment_test_users()
