def untag_candidate(self, token):
    data1 = [{"testId": 9214, "candidateIds": [1292529, 1292533, 1292530]},
             {"testId": 9216, "candidateIds": [1292529]},
             {"testId": 9218, "candidateIds": [1292529]},
             {"testId": 9220, "candidateIds": [1292534, 1292533, 1292530, 1292529]}]
    for request in data1:
        print(request)
        response = requests.post("https://amsin.hirepro.in/py/assessment/testuser/api/v1/un-tag/",
                                 headers=token,
                                 data=json.dumps(request, default=str), verify=False)