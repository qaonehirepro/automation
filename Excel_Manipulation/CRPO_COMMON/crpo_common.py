import requests
import json


class CrpoCommon:
    domain = 'https://amsin.hirepro.in'

    @staticmethod
    def login_to_crpo(login_name, password, tenant):
        header = {"content-type": "application/json"}
        data = {"LoginName": login_name, "Password": password, "TenantAlias": tenant, "UserName": login_name}
        response = requests.post(crpo_common_obj.domain + "/py/common/user/login_user/", headers=header,
                                 data=json.dumps(data), verify=False)
        login_response = response.json()
        headers = {"content-type": "application/json", "APP-NAME": "py3app", "X-APPLMA": "true",
                   "X-AUTH-TOKEN": login_response.get("Token")}
        return headers

    @staticmethod
    def candidate_web_transcript(token, test_id, test_user_id):
        request = {"testId": int(test_id), "testUserId": int(test_user_id),
                   "reportFlags": {"eduWorkProfilesRequired": True, "testUsersScoreRequired": True,
                                   "fileContentRequired": False, "isProctroingDetailsRequired": True}, "print": False}
        response = requests.post(crpo_common_obj.domain + "/py/assessment/report/api/v1/candidatetranscript/",
                                 headers=token,
                                 data=json.dumps(request, default=str), verify=False)
        return response.json()

    @staticmethod
    def force_evaluate_proctoring(token, tu_ids):
        print(token)
        request = {
            "testUserIds": tu_ids, "isForce": True}
        response = requests.post(crpo_common_obj.domain +
                                 "/py/assessment/htmltest/api/v1/initiate-test-proc/?isSync=false",
                                 headers=token,
                                 data=json.dumps(request, default=str), verify=False)
        return response.json()

    @staticmethod
    def job_status(token, contextguid):
        print(token)
        request = {"ContextGUID": contextguid}
        response = requests.post(crpo_common_obj.domain + "/py/crpo/api/v1/getStatusOfAsyncAPI",
                                 headers=token, data=json.dumps(request, default=str), verify=False)
        resp_dict = json.loads(response.content)
        # api_job_status = resp_dict['data']['JobState']
        # return api_job_status
        return resp_dict

    @staticmethod
    def upload_files(token, file_name, file_path):
        # This API does not require content type and its not Lambda API as of 29/04/2021
        # so poping the values from token.
        token.pop('content-type', None)
        token.pop('X-APPLMA', None)
        request = {'file': (file_name, open(file_path, 'rb'))}
        url = crpo_common_obj.domain + '/py/common/filehandler/api/v2/upload/.doc,.rtf,.dot,.docx,' \
                                       '.docm,.dotx,.dotm,.docb,.pdf,.xls,.xlt,.xlm,.xlsx,.xlsm,.xltx,.xltm,.xlsb,.xla,.xlam,.xll,' \
                                       '.xlw,.ppt,.pot,.pps,.pptx,.pptm,.potx,.potm,.ppam,.ppsx,.ppsm,.sldx,.sldm,.zip,.rar,.7z,.gz,.jpeg,' \
                                       '.jpg,.gif,.png,.msg,.txt,.mp4,.mvw,.3gp,.sql,.webm,.csv,.odt,.json,.ods,.ogg,.p12,/5000/'

        api_request = requests.post(url, headers=token, files=request, verify=False)
        # print(api_request.headers.get('X-GUID'))
        resp_dict = json.loads(api_request.content)
        return resp_dict

    @staticmethod
    def untag_candidate(token, data1):
        # sample data = [{"testUserIds": [893441, 893442, 893443]}]
        for request in data1:
            response = requests.post(crpo_common_obj.domain + "/py/assessment/testuser/api/v1/un-tag/",
                                     headers=token,
                                     data=json.dumps(request, default=str), verify=False)

    @staticmethod
    def proctor_evaluation_detail(token, testuser_id):
        token.pop('X-APPLMA', None)
        print(token)
        request = {"testUserId": testuser_id}
        response = requests.post(crpo_common_obj.domain + "/py/assessment/testuser/api/v1/get_proctor_detail/",
                                 headers=token,
                                 data=json.dumps(request, default=str), verify=False)
        tu_proctor_details = response.json()
        return tu_proctor_details

    @staticmethod
    def re_initiate_automation(token, test_id, candidate_id):
        token.pop('X-APPLMA', None)
        print(token)
        request = {"testId": test_id, "candidateId": candidate_id}
        response = requests.post(crpo_common_obj.domain + "/py/assessment/testuser/api/v1/re_initiate_automation/",
                                 headers=token,
                                 data=json.dumps(request, default=str), verify=False)

    @staticmethod
    def get_all_questions(token, request_data):
        response = requests.post(crpo_common_obj.domain + "/py/assessment/authoring/api/v1/getAllQuestion/",
                                 headers=token,
                                 data=str(request_data.get('request')), verify=False)
        get_all_questions_resp = json.loads(response.content)
        return get_all_questions_resp

    @staticmethod
    def generate_applicant_report(token, request_payload):
        response = requests.post(crpo_common_obj.domain + "/py/common/xl_creator/api/v1/generate_applicant_report/",
                                 headers=token, data=json.dumps(request_payload, default=str), verify=False)
        resp_dict = json.loads(response.content)
        return resp_dict

    @staticmethod
    def generate_plagiarism_report(token, request_payload):
        response = requests.post(crpo_common_obj.domain + "/py/assessment/report/api/v1/plagiarismreport/",
                                 headers=token, data=json.dumps(request_payload, default=str), verify=False)
        resp_dict = json.loads(response.content)
        return resp_dict

    @staticmethod
    def initiate_vendor_score(crpotoken, cid, test_id):
        url = crpo_common_obj.domain + '/py/assessment/assessmentvendor/api/v1/initiateVendorScore/'
        data = {"testId": test_id, "candidateIds": [cid], "isForced": True}

        response = requests.post(url,
                                 headers=crpotoken,
                                 data=json.dumps(data, default=str), verify=False)
        it_vendor_resp = response.json()
        print(it_vendor_resp)
        return it_vendor_resp

    @staticmethod
    def untag_candidate_by_cid(token, test_id, candidate_ids):
        data1 = [{"testId": test_id, "candidateIds": candidate_ids}]
        for request in data1:
            response = requests.post(crpo_common_obj.domain + "/py/assessment/testuser/api/v1/un-tag/",
                                     headers=token,
                                     data=json.dumps(request, default=str), verify=False)
            print(response)

    @staticmethod
    def create_candidate(token, usn):
        request = {"PersonalDetails": {"FirstName": usn, "Email1": usn + "qaone.h.i.repro@gmail.com", "USN": usn}}
        # req = {"PersonalDetails": {"FirstName": "MuthuMurugan", "Email1": "qaonehirepro@gmail.com", "USN": "151_2"}}
        response = requests.post('https://amsin.hirepro.in/py/rpo/create_candidate/', headers=token,
                                 data=json.dumps(request), verify=False)
        response_data = response.json()
        candidate_id = response_data.get('CandidateId')
        if response_data.get('status') == 'OK':
            print("candidate created in crpo")
            url = 'https://automation-in.hirepro.in/?candidate=%s' % candidate_id
        else:
            print("candidate not created in CRPO_COMMON due to some technical glitch")
            print(response_data)
        return candidate_id

    @staticmethod
    def tag_candidate_to_test(token, cid, testid, eventid, jobroleid):
        request = {"CandidateIds": [int(cid)], "TestIds": [int(testid)], "EventId": int(eventid),
                   "JobRoleId": int(jobroleid), "Sync": "True"}
        response = requests.post(
            "https://amsin.hirepro.in/py/crpo/applicant/api/v1/tagCandidatesToEventJobRoleTests/",
            headers=token,
            data=json.dumps(request, default=str), verify=False)
        return response

    @staticmethod
    def test_user_credentials(token, tu_id):
        request = {"testUserId": tu_id}
        response = requests.post(
            "https://amsin.hirepro.in/py/assessment/testuser/api/v1/getCredential/",
            headers=token,
            data=json.dumps(request, default=str), verify=False)
        # data = response.json()
        return response.json()

    @staticmethod
    def get_all_test_user(token, cid):
        request = {"isMyAssessments": False, "search": {"candidateIds": [cid]}}
        response = requests.post("https://amsin.hirepro.in/py/assessment/testuser/api/v1/getAllTestUser/",
                                 headers=token,
                                 data=json.dumps(request, default=str), verify=False)
        data = response.json()
        test_user_id = data['data']['testUserInfos'][0]['id']
        return test_user_id


crpo_common_obj = CrpoCommon()
