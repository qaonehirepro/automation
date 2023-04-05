import json

automation_proctor_eval_app_pref = {
    "video": {"weightage": 0.3, "zero_count": 3, "multiple_count": 1, "zero_timespan": 10,
              "multipleIntervalCount": {"isEnabled": True,
                                        "intervalCount": [{"weightage": 0.32, "count": 1, "interval": 60},
                                                          {"weightage": 0.31, "count": 2, "interval": 30}]},
              "isEnabled": True, "multiple_timespan": 10, "zeroIntervalCount": {"isEnabled": True, "intervalCount": [
            {"weightage": 0.34, "count": 1, "interval": 60}, {"weightage": 0.33, "count": 2, "interval": 30}]}},
    "audio": {"weightage": 0.2, "isEnabled": True, "no_of_words": 10}, "overall": 0.1,
    "face": {"faceMisMatch": {"weightage": 0.6, "suspiciousThreshold": 6, "isEnabled": True},
             "noFace": {"weightage": 0.5, "suspiciousThreshold": 10, "isEnabled": True},
             "goldenImage": {"weightage": 0.9, "isEnabled": True},
             "suspiciousCount": {"weightage": 0.8, "isEnabled": True, "suspiciousThresholdPercentage": 80},
             "isEnabled": True, "multipleFace": {"weightage": 0.4, "suspiciousThreshold": 10, "isEnabled": True}},
    "device": {"isEnabled": True,
               "computer": {"isEnabled": True, "multipleCamera": {"isEnabled": True, "count": 2, "weightage": 0.555},
                            "singleCamera": {"isEnabled": True, "weightage": 0.222, "keywords": ["Integrated", "HP"],
                                             "SuspiciousCamera": {"isEnabled": True, "keywords": ["Facetime", "Logi"],
                                                                  "weightage": 0.888}}}}}

print(json.dumps(automation_proctor_eval_app_pref))
