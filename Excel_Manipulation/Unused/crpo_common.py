import requests
import json


class CrpoCommon:

    @staticmethod
    def untag_candidate(token, test_id, candidate_ids):
        data1 = [{"testId": test_id, "candidateIds": candidate_ids}]
        for request in data1:
            response = requests.post("https://amsin.hirepro.in/py/assessment/testuser/api/v1/un-tag/",
                                     headers=token,
                                     data=json.dumps(request, default=str), verify=False)
            print(response)

    # @staticmethod
    # def login_to_crpo(login_name, password, tenant):
    #     header = {"content-type": "application/json"}
    #     data = {"LoginName": login_name, "Password": password, "TenantAlias": tenant, "UserName": login_name}
    #     response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
    #                              data=json.dumps(data), verify=False)
    #     login_response = response.json()
    #     headers = {"content-type": "application/json", "X-AUTH-TOKEN": login_response.get("Token")}
    #     return headers

crpo_common_obj1 = CrpoCommon()