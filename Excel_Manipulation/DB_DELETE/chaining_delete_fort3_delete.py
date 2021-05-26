import mysql
import mysql.connector
import datetime


class delete_ssrf_data:

    def __init__(self):
        print(datetime.datetime.now())

    def amsdbconnection(self):
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='qauser',
                                            password='qauser')
        self.cursor = self.conn.cursor()

    def commit_changes(self, query):
        pass

    def delete_assessment_test_users(self):
        self.amsdbconnection()
        tuser_scores = 'delete from candidate_scores where testuser_id in ' \
                       '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                       'where test_id in(10508,10509,10510,10568,10569,10570,10581,10582,10584,10698,10699,10700,10709,10710,10711,10713,10714,10715)' \
                       ' and login_time is not null and t.tenant_id in (159,1786));'
        print(tuser_scores)

        self.cursor.execute(tuser_scores)
        self.conn.commit()
        tuser_login_infos = 'delete from test_user_login_infos where testuser_id in ' \
                            '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                            'where test_id in(10508,10509,10510,10568,10569,10570,10581,10582,10584,10698,10699,10700,10709,10710,10711,10713,10714,10715) ' \
                            'and login_time is not null and t.tenant_id in (159,1786));'
        print(tuser_login_infos)
        self.cursor.execute(tuser_login_infos)
        self.conn.commit()

        tuser_proctoring_infos = 'delete from test_user_proctor_details where testuser_id in ' \
                                 '(select tu.id from test_users tu inner join tests t on t.id = tu.test_id ' \
                                 'where test_id in(10508,10509,10510,10568,10569,10570,10581,10582,10584,10698,10699,10700,10709,10710,10711,10713,10714,10715) ' \
                                 'and login_time is not null and t.tenant_id in (159,1786));'
        print(tuser_proctoring_infos)
        self.cursor.execute(tuser_proctoring_infos)
        self.conn.commit()

        update_tuser_statuss = 'update test_users set login_time = NULL, is_disabled = 0, log_out_time = NULL, ' \
                               'status = 0, client_system_info = NULL, time_spent = NULL, is_password_disabled = 0,' \
                               'config = NULL, client_system_info = NULL, total_score = NULL, percentage = NULL ' \
                               'where test_id in(10508,10509,10510,10568,10569,10570,10581,10582,10584,10698,10699,10700,10709,10710,10711,10713,10714,10715) and login_time is not null;'
        print(update_tuser_statuss)
        self.cursor.execute(update_tuser_statuss)
        self.conn.commit()
        # # delete_question_approval = 'delete from question_approvals where question_id =\'113596\' and tenant_id=159;'
        # # self.cursor.execute(delete_question_approval)
        # # self.conn.commit()
        """# don't add tu id for cocubes, mettl, wheebox...etc  Vendors her"""
        update_test_users_partner_infos = 'update test_users_partner_info set status=3, partner_uuid = NULL, ' \
                                          'remote_candidate_json= NULL, score_status = NULL,task_id_score_fetch = NULL,' \
                                          ' report_link = NULL, tenant_id = NULL, third_party_status = NULL,' \
                                          ' third_party_login_time = NULL,third_party_test_link = NULL, ' \
                                          'third_party_overall_status = NULL, communication_history_json = NULL  ' \
                                          'where testuser_id in (885141,885142,885143,885144,885145,885165,885166,885167,885168,885169);'

        print(update_test_users_partner_infos)
        self.cursor.execute(update_test_users_partner_infos)
        self.conn.commit()
        # """ add tu id for cocubes, mettl, wheebox...etc  Vendors here."""
        # update_test_users_partner_info_for_pull_score = 'update  test_users_partner_info set score_status = Null, ' \
        #                                                 'task_id_score_fetch = Null, communication_history_json =Null, ' \
        #                                                 'report_link = Null, third_party_status =Null, ' \
        #                                                 'third_party_overall_status = Null where ' \
        #                                                 'testuser_id in (884052);'
        # print(update_test_users_partner_info_for_pull_score)
        # self.cursor.execute(update_test_users_partner_info_for_pull_score)
        # self.conn.commit()
        self.conn.close()


del_data = delete_ssrf_data()
del_data.delete_assessment_test_users()

print(datetime.datetime.now())
