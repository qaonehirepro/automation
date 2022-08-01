qp_qn_index = [{"qn": "MS Question Randomization Low 1", "index": 1},
               {"qn": "MS Question Randomization Low 2", "index": 2},
               {"qn": "MS Question Randomization Low 3", "index": 3},
               {"qn": "MS Question Randomization medium11", "index": 4},
               {"qn": "MS Question Randomization medium21", "index": 5},
               {"qn": "MS Question Randomization medium31", "index": 6},
               {"qn": "MS Question Randomization high1", "index": 7},
               {"qn": "MS Question Randomization high2", "index": 8},
               {"qn": "MS Question Randomization high3", "index": 9},
               {"qn": "MS Question Randomization Low 4", "index": 10},
               {"qn": "MS Question Randomization Low 5", "index": 11},
               {"qn": "MS Question Randomization medium4", "index": 12},
               {"qn": "MS Question Randomization medium5", "index": 13},
               {"qn": "MS Question Randomization high4", "index": 14},
               {"qn": "MS Question Randomization high5", "index": 15},
               {"qn": "MS Question Randomization Low 6", "index": 16},
               {"qn": "MS Question Randomization Low 7", "index": 17},
               {"qn": "MS Question Randomization high6", "index": 18},
               {"qn": "MS Question Randohighmization 7", "index": 19},
               {"qn": "MS Question Randomization medium6", "index": 20},
               {"qn": "MS Question Randomization medium7", "index": 21}]

section1_group1_questions = ["MS Question Randomization Low 1", "MS Question Randomization Low 2",
                                  "MS Question Randomization Low 3", "MS Question Randomization medium11",
                                  "MS Question Randomization medium21", "MS Question Randomization medium31",
                                  "MS Question Randomization high1", "MS Question Randomization high2",
                                  "MS Question Randomization high3"]
section2_group1_questions = ["MS Question Randomization Low 4", "MS Question Randomization Low 5",
                                  "MS Question Randomization medium4", "MS Question Randomization medium5",
                                  "MS Question Randomization high4", "MS Question Randomization high5"]

section1_group2_questions = ["MS Question Randomization Low 1", "MS Question Randomization Low 7",
                                  "MS Question Randomization high6", "MS Question Randomization high7",
                                  "MS Question Randomization medium6", "MS Question Randomization medium7"]

qn_details = {'question': "MS Question Randomization Low 7", 'group': "Group1", 'section': "Section1",
                                       'index': 1}


def is_randomized(qn_details):
    a = 1
    for all_qns_with_index in qp_qn_index:
        a += 1
        print(a)
        print("________________________________________________________\n")
        if qn_details.get('question') == all_qns_with_index.get("qn"):

            if qn_details.get('index') != all_qns_with_index.get("index"):
                overall_randomization = "Yes"
                if 0 <= qn_details.get('index') <= 14:
                    if qn_details.get('question') in section1_group1_questions:
                        g1_randomization = "Yes"
                        print("This is muthu1", a)
                elif 15 <= all_qns_with_index.get("index") <= 20:
                    if qn_details.get('question') in section1_group2_questions:
                        g2_randomization = "Yes"
                        print("This is muthu2", a)
            break



is_randomized(qn_details)