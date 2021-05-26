import requests
import json
from ASSESSMENT import submit_test_data

requests.packages.urllib3.disable_warnings()


class AssessmentCommonMethods:

    def decide_domain(self, type_of_test):
        print(type_of_test)
        if type_of_test == 'HP':
            domain = "https://amsin.hirepro.in/py/assessment/"
        elif type_of_test == 'Cocubes':
            domain = "https://amsin.hirepro.in/py/assessment/"
        elif type_of_test == 'TALENTLENS':
            domain = "https://talentlensstg.hirepro.in/py/assessment/"
        elif type_of_test == 'VET':
            domain = "https://pearsonstg.hirepro.in/py/assessment/"
        else:
            domain = "https://amsin.hirepro.in/py/assessment/"

        return domain


    @staticmethod
    def login_to_test(login_name, password, tenant, type_of_test):
        print(login_name, password, type_of_test)
        if type_of_test == 'HP':
            domain = "https://amsin.hirepro.in/py/assessment/"
        elif type_of_test == 'Cocubes':
            domain = "https://amsin.hirepro.in/py/assessment/"
        elif type_of_test == 'TALENTLENS':
            domain = "https://talentlensstg.hirepro.in/py/assessment/"
        elif type_of_test == 'VET':
            domain = "https://pearsonstg.hirepro.in/py/assessment/"
        else:
            domain = "https://amsin.hirepro.in/py/assessment/"

        header = {"content-type": "application/json"}
        data = {"LoginName": login_name, "Password": password, "TenantAlias": tenant}
        login_url = domain + 'htmltest/api/v2/login_to_test/'
        response = requests.post(login_url, headers=header, data=json.dumps(data), verify=False)
        login_response = response.json()
        login_token = {}
        test_user_infos = {}
        if login_response.get('status') == 'OK':
            test_type = json.loads(login_response.get('Config')).get('thirdPartyTestType')
            if not test_type:
                test_type = 'HP'
            are_you_able_to_login = 'Yes'
            login_token = {'X-AUTH-TOKEN': login_response.get("Token")}
            test_user_infos = {'nextTestID': login_response.get("TestId"),
                               'nextTestName': login_response.get("TestName"),
                               'nextTestCandidateId': login_response.get("CandidateId"),
                               'login_staus': are_you_able_to_login, 'next_test_type': test_type}

        elif login_response.get('error').get('errorCode'):
            nextTestFlags = login_response.get('error').get('nextTestFlags')
            if nextTestFlags.get('isShortlisted') or nextTestFlags.get('isShortlisted') == None:
                are_you_able_to_login = 'Yes'
            else:
                are_you_able_to_login = 'No'
            print("Login Failed")
            test_type = 'EMPTY'
            test_user_infos = {'nextTestID': "Login Failed",
                               'nextTestName': "Login Failed",
                               'nextTestCandidateId': "Login Failed",
                               'login_staus': are_you_able_to_login, 'next_test_type': test_type}

        else:
            print("Login Failed")
            are_you_able_to_login = 'No'
            test_type = 'EMPTY'
            test_user_infos = {'nextTestID': "Login Failed",
                               'nextTestName': "Login Failed",
                               'nextTestCandidateId': "Login Failed",
                               'login_staus': are_you_able_to_login, 'next_test_type': test_type}

        login_resp = {"login_token": login_token, "domain": domain, "test_user_infos": test_user_infos,
                      "login_response": login_response}

        return login_resp

    @staticmethod
    def submit_test_result(assessment_token, domain, submit_test_request):

        url = domain + 'htmltest/api/v1/finalSubmitTestResult/'
        data1 = submit_test_data.alldata.get(str(submit_test_request))
        response = requests.post(url,
                                 headers=assessment_token,
                                 data=json.dumps(data1, default=str), verify=False)
        json_resp = response.json()
        if json_resp.get('isResultSubmitted'):
            submit_token = json_resp.get('systemTkn')
        else:
            print("The Result is not submitted")
            submit_token = None
        submit_token = {'X-AUTH-TOKEN': submit_token}
        return submit_token

    @staticmethod
    def initiate_automation(submit_token, cid, test_id, domain):
        context_id = None
        next_test_info = None
        # url = domain + 'testuser/api/v1/initiate_automation/'
        url = domain + 'testuser/api/v2/initiate_automation/'
        data = {"candidateId": int(cid), "testId": test_id, "debugTimeStamp": "2020-07-14T07:32:54.904Z"}
        response = requests.post(url,
                                 headers=submit_token,
                                 data=json.dumps(data, default=str), verify=False)
        itua_resp = response.json()
        return itua_resp

    @staticmethod
    def get_job_status(token, context_id):
        json_resp = {}
        status = 'Pending'
        data = {"ContextGUID": context_id, "disableBlockUI": True,
                "debugTimeStamp": "2020-07-15T07:49:04.013Z"}
        while status != 'SUCCESS':
            response = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/getStatusOfAsyncAPI",
                                     headers=token,
                                     data=json.dumps(data, default=str), verify=False)
            json_resp = response.json()
            status = json_resp['data']['JobState']
            print(json_resp)

        return json_resp

    # def find_next_test_status_for_chaining(self, login_response):
    #     relogin_status = login_response
    #     if relogin_status.get('status') == 'KO' and relogin_status.get('statusCode') == 200:
    #         print("You have already submitted")
    #     elif relogin_status.get('status') == 'KO' and relogin_status.get('statusCode') == 500:
    #         if relogin_status['error']['nextTestFlags']["isShortlisted"]:
    #             print("candidate is shortlisted for next test")
    #             print("Need to call next test api here.")
    #             # next_test_result = Next_test['nextTestInfos'][0]
    #             # relogin_t2_testid = next_test_result.get('testId')
    #             # relogin_t2_login_name = next_test_result.get('loginId')
    #             # relogin_t2_password = next_test_result.get('password')
    #             # print(relogin_t2_testid, relogin_t2_login_name, relogin_t2_password)
    #         else:
    #             print("candidate is rejected for next test")
    #             print("No Need to call next test api here.")
    #     else:
    #         print(" Candidate is not completed the first test ")

    @staticmethod
    def process_next_test_links_for_chaining(json_resp, previous_domain):
        print("job status changed to success")
        job_status_result = json_resp['data']['Result']
        next_test_info = json.loads(job_status_result)
        is_shortlisted = next_test_info['data']['isAutoShortlisting']
        if is_shortlisted:
            print("This is SLC test")
            actual_status = "Shortlisted"
        else:
            print("This is Auto Test")
            actual_status = "EMPTY"
        if next_test_info['data']['nextTestInfos']:
            is_next_test_infos_available = "Yes"
            next_test_infos = (next_test_info['data']['nextTestInfos'][0])
            proxy_url = next_test_infos.get('proxyUrl')
            if proxy_url:
                domain = proxy_url.split('/')
                domain = domain[2]
                host = previous_domain.split('/')
                host = host[1]
                next_test_domain_and_host = domain + '/' + host
            else:
                proxy_url = previous_domain
                next_test_domain_and_host = previous_domain
            next_test_login_id = next_test_infos.get('loginId')
            next_test_pwd = next_test_infos.get('password')

        else:
            actual_status = "Rejected"
            next_test_domain_and_host = previous_domain
            is_next_test_infos_available = "No"
            next_test_domain_and_host = "EMPTY"
            next_test_login_id = "EMPTY"
            next_test_pwd = "EMPTY"
            print("Candidate is not shortlisted for the next test due to low score or no education profile.")

        next_test = {'actualStatus': actual_status, "nextTestDomainHost": next_test_domain_and_host,
                     'nextTestAvailability': is_next_test_infos_available,
                     'next_test_login_id': next_test_login_id,
                     'next_test_pwd': next_test_pwd}
        return next_test

    @staticmethod
    def pearson_call_backs(test_user_id, score_details):
        content = """ <?xml version = "1.0" encoding = "UTF-8"?>
        <imsx_POXEnvelopeRequest xmlns = "http://www.imsglobal.org/services/ltiv1p1/xsd/imsoms_v1p0">
          <imsx_POXHeader>
            <imsx_POXRequestHeaderInfo>
              <imsx_version>V1.0</imsx_version>
              <imsx_messageIdentifier>7777897</imsx_messageIdentifier>
            </imsx_POXRequestHeaderInfo>
          </imsx_POXHeader>	
          <imsx_POXBody>
            <replaceResultRequest>
              <resultRecord>
                <sourcedGUID>
                  <sourcedId>{test_user_id}</sourcedId>
                </sourcedGUID>
                <tin>
                  <tinID>28035214</tinID>
                </tin>		
                <result>
                  <resultScore>
                    <language>en</language>						
                    <textString>{score_response}</textString>
                  </resultScore>
                </result>
              </resultRecord>
            </replaceResultRequest>
          </imsx_POXBody>
        </imsx_POXEnvelopeRequest> """

        content = content.format(test_user_id="%s" % test_user_id, score_response="%s" % score_details)
        response = requests.post(
            "https://amsin.hirepro.in/py/assessment/assessmentvendor/api/v1/vcb/versant/?tn=76EF28AF-6DB5-11EA-8197-0262BDD19558",
            headers={"Content-Type": "application/xml"}, data=content)
        print(response)
        print(response.content)

    @staticmethod
    def initiate_vendor_score(crpotoken, cid, test_id):
        url = 'https://amsin.hirepro.in/py/assessment/assessmentvendor/api/v1/initiateVendorScore/'
        data = {"testId": test_id, "candidateIds": [cid], "isForced": True}

        response = requests.post(url,
                                 headers=crpotoken,
                                 data=json.dumps(data, default=str), verify=False)
        it_vendor_resp = response.json()
        return it_vendor_resp

    @staticmethod
    def next_test_info_for_2nd_login(login_name, password, tenant, domain):
        second_login_data = {}
        header = {"content-type": "application/json"}
        data = {"loginName": login_name, "password": password, "tenantAlias": tenant,
                "debugTimeStamp": "2020-12-02T13:32:30.749Z"}
        login_url = domain + 'htmltest/api/v1/test-user-next_test/'
        response = requests.post(login_url, headers=header, data=json.dumps(data), verify=False)
        second_login_response = response.json()
        data = second_login_response.get('nextTestInfos')[0]
        if data.get('isShortlisted'):
            second_login_is_shortlisted = "slc"
        else:
            second_login_is_shortlisted = "auto"
        second_login_data = {'second_login_cid': data.get('candidateId'), 'second_login_test_id': data.get('testId'),
                             'second_test_login_id': data.get('loginId'), 'second_test_password': data.get('password'),
                             'login_url': domain, 'second_login_is_shortlisted': second_login_is_shortlisted}
        return second_login_data


assessment_common_obj = AssessmentCommonMethods()
