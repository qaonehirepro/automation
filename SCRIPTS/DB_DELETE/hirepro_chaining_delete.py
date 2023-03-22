from SCRIPTS.COMMON.dbconnection import *
import datetime


class DeleteHireproChaining:

    def __init__(self):
        print(datetime.datetime.now())

    def delete_assessment_test_users(self):
        db_connection = ams_db_connection()
        cursor = db_connection.cursor()
        tuser_scores = 'delete from candidate_scores where testuser_id in ' \
                       '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                       'where test_id in(12397,12399,12401,12403,12405,12407)' \
                       ' and login_time is not null and t.tenant_id in (1787));'
        print(tuser_scores)

        cursor.execute(tuser_scores)
        # self.conn.commit()
        tuser_login_infos = 'delete from test_user_login_infos where testuser_id in ' \
                            '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                            'where test_id in(12397,12399,12401,12403,12405,12407) ' \
                            'and login_time is not null and t.tenant_id in (1787));'
        print(tuser_login_infos)
        cursor.execute(tuser_login_infos)
        # self.conn.commit()

        tuser_proctoring_infos = 'delete from test_user_proctor_details where testuser_id in ' \
                                 '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                                 'where test_id in(12397,12399,12401,12403,12405,12407) ' \
                                 'and login_time is not null and t.tenant_id in (1787));'
        print(tuser_proctoring_infos)
        cursor.execute(tuser_proctoring_infos)
        # self.conn.commit()

        update_tuser_statuss = 'update test_users set login_time = NULL, log_out_time = NULL, status = 0, ' \
                               'client_system_info = NULL, time_spent = NULL, is_password_disabled = 0,config = NULL, ' \
                               'client_system_info = NULL, total_score = NULL, percentage = NULL ' \
                               'where test_id in(12397,12399,12401,12403,12405,12407) and ' \
                               'login_time is not null;'
        print(update_tuser_statuss)
        cursor.execute(update_tuser_statuss)
        # self.conn.commit()
        # delete_question_approval = 'delete from question_approvals where question_id =\'113596\' and tenant_id=159;'
        # self.cursor.execute(delete_question_approval)
        # self.conn.commit()
        """# don't add tu id for cocubes, mettl, wheebox...etc  Vendors test"""
        # update_test_users_partner_infos = 'update test_users_partner_info set status=3, partner_uuid = NULL, ' \
        #                                   'remote_candidate_json= NULL, score_status = NULL,task_id_score_fetch = NULL,' \
        #                                   ' report_link = NULL, tenant_id = NULL, third_party_status = NULL,' \
        #                                   ' third_party_login_time = NULL,third_party_test_link = NULL, ' \
        #                                   'third_party_overall_status = NULL, communication_history_json = NULL  ' \
        #                                   'where testuser_id in (880531,880555,880556,880557,880558,880559,880561,' \
        #                                   '880562,880563,880564,880565,880594,880596,880598,882985,882984,882983,' \
        #                                   '882982,882988,882989,882990,882991);'

        # print(update_test_users_partner_infos)
        # self.cursor.execute(update_test_users_partner_infos)
        # self.conn.commit()
        """ add tu id for cocubes, mettl, wheebox...etc  Vendors here."""
        # update_test_users_partner_info_for_pull_score = 'update  test_users_partner_info set score_status = Null, ' \
        #                                                 'task_id_score_fetch = Null, communication_history_json =Null, ' \
        #                                                 'report_link = Null, third_party_status =Null, ' \
        #                                                 'third_party_overall_status = Null where ' \
        #                                                 'testuser_id in (882370,882393,882400,882401,882402,882484,' \
        #                                                 '882485,882486,882487,882505,882506,882507,882508);'
        # print(update_test_users_partner_info_for_pull_score)
        # self.cursor.execute(update_test_users_partner_info_for_pull_score)
        db_connection.commit()
        db_connection.close()


del_data = DeleteHireproChaining()
del_data.delete_assessment_test_users()
print(datetime.datetime.now())
