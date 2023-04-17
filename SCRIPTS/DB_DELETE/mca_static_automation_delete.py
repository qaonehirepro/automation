from SCRIPTS.COMMON.dbconnection import *
import datetime


class UIAutomationDelete:

    def __init__(self):
        print(datetime.datetime.now())

    def delete_assessment_test_users(self):
        db_connection = ams_db_connection()
        cursor = db_connection.cursor()
        tuser_results = "delete from test_results where testuser_id in (2551271,2551269,2551267,2551265,2551263);"
        cursor.execute(tuser_results)
        db_connection.commit()
        tuser_scores = 'delete from candidate_scores where testuser_id in ' \
                       '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                       'where test_id in(16287)' \
                       ' and login_time is not null and t.tenant_id in (159,1786));'
        print(tuser_scores)

        cursor.execute(tuser_scores)
        db_connection.commit()
        tuser_login_infos = 'delete from test_user_login_infos where testuser_id in ' \
                            '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                            'where test_id in(16287) ' \
                            'and login_time is not null and t.tenant_id in (159,1786));'
        print(tuser_login_infos)
        cursor.execute(tuser_login_infos)
        db_connection.commit()

        tuser_proctoring_infos = 'delete from test_user_proctor_details where testuser_id in ' \
                                 '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                                 'where test_id in(16287) ' \
                                 'and login_time is not null and t.tenant_id in (159,1786));'
        print(tuser_proctoring_infos)
        cursor.execute(tuser_proctoring_infos)
        db_connection.commit()

        update_tuser_statuss = 'update test_users set login_time = NULL, log_out_time = NULL, status = 0, ' \
                               'client_system_info = NULL, time_spent = NULL, is_password_disabled = 0,config = NULL, ' \
                               'client_system_info = NULL, total_score = NULL, test_start_time = NULL, percentage = NULL, ' \
                               'correct_answers = NULL, in_correct_answers = NULL, un_attended_questions=NULL,'\
                               ' is_partially_evaluated = NULL, eval_status = "NotEvaluated", eval_on = NULL'\
                               ' where test_id in(16287) and ' \
                               'login_time is not null;'
        print(update_tuser_statuss)
        cursor.execute(update_tuser_statuss)
        db_connection.commit()
        db_connection.close()


del_data = UIAutomationDelete()
del_data.delete_assessment_test_users()
print(datetime.datetime.now())
