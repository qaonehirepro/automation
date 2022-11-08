import json

#
# parent = "MS RTC Serverside Parent Qn2"
# child = "c22"
# all_questions = [{"parent": "MS RTC Serverside Parent Qn1",
#                   "childs": [{"child": "c11"}, {"child": "c12"}, {"child": "c13"}]},
#                  {"parent": "MS RTC Serverside Parent Qn2",
#                   "childs": [{"child": "c21"}, {"child": "c22"}, {"child": "c23"}]},
#                  {"parent": "MS RTC Serverside Parent Qn3",
#                   "childs": [{"child": "c31"}, {"child": "c32"}, {"child": "c33"}]}]
#
# for question in all_questions:
#     if parent in question.get("parent"):
#         expected_parent = "Yes"
#         childs = question.get('childs')
#         for childsforquestion in childs:
#             if child in childsforquestion.get('child'):
#                 expected_child = "Yes"


# payload = {"testUserId": 2452472,
#                    "requiredFlags": {"fileContentRequired": False, "isQuestionWise": False, "questionTypes": [16, 8],
#                                      "isGroupSectionWiseMarks": True, "isVendorDetails": False,
#                                      "isCodingSummary": False}}
# a = 12345
# payload['testUserId'] = a
# print(payload)

mark_details = {
    "status": "OK",
    "statusId": 200,
    "data": {
        "candidate": {
            "id": 1412954,
            "firstName": "Ms",
            "lastName": "Ld 5000 Aug22 0982",
            "candidateName": "Ms Cd Ld 5000 Aug22 0982",
            "email1": "qaone.hirepro@gmail.com",
            "email2": None,
            "mobile1": "9842321601",
            "middleName": "Cd",
            "address1": None,
            "address2": None,
            "address3": None,
            "dateOfBirth": None,
            "totalExperience": 0,
            "currentExperience": None,
            "currentLocationId": None,
            "currentLocationText": None,
            "photoUrl": None
        },
        "educationProfiles": [],
        "workProfiles": [],
        "testResultQuestionIds": [
            121400,
            121402,
            121404,
            121410,
            121412,
            121414,
            121420,
            121422,
            121424,
            121406,
            121408,
            121416,
            121418,
            121426,
            121428,
            121442,
            121444,
            121462,
            121464,
            121484,
            121486
        ],
        "mcq": [
            {
                "parentQuestionId": None,
                "id": 121400,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 1",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 1",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.1,
                "correctAnswer": "A",
                "timeSpent": 7,
                "candidateAnswer": "B",
                "obtainedMark": -0.1
            },
            {
                "parentQuestionId": None,
                "id": 121402,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 2",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 2",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.2,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.2
            },
            {
                "parentQuestionId": None,
                "id": 121404,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 3",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 3",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.3,
                "correctAnswer": "A",
                "timeSpent": 2,
                "candidateAnswer": "B",
                "obtainedMark": -0.3
            },
            {
                "parentQuestionId": None,
                "id": 121410,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium11",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium11",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.4,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.4
            },
            {
                "parentQuestionId": None,
                "id": 121412,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium21",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium21",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.5,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.5
            },
            {
                "parentQuestionId": None,
                "id": 121414,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium31",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium31",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.6,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.6
            },
            {
                "parentQuestionId": None,
                "id": 121420,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high1",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high1",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.7,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.7
            },
            {
                "parentQuestionId": None,
                "id": 121422,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high2",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high2",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.8,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.8
            },
            {
                "parentQuestionId": None,
                "id": 121424,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high3",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64616,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high3",
                "sectionName": "Group1Section 1",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 0.9,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -0.9
            },
            {
                "parentQuestionId": None,
                "id": 121406,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 4",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64618,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 4",
                "sectionName": "Group1Section2",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1
            },
            {
                "parentQuestionId": None,
                "id": 121408,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 5",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64618,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 5",
                "sectionName": "Group1Section2",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 1.1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1.1
            },
            {
                "parentQuestionId": None,
                "id": 121416,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium4",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64618,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium4",
                "sectionName": "Group1Section2",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 1.2,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1.2
            },
            {
                "parentQuestionId": None,
                "id": 121418,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium5",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64618,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium5",
                "sectionName": "Group1Section2",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 1.3,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1.3
            },
            {
                "parentQuestionId": None,
                "id": 121426,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high4",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64618,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high4",
                "sectionName": "Group1Section2",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 1.4,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1.4
            },
            {
                "parentQuestionId": None,
                "id": 121428,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high5",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64618,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high5",
                "sectionName": "Group1Section2",
                "groupName": "Group1",
                "groupId": 25490,
                "mark": 1.5,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1.5
            },
            {
                "parentQuestionId": None,
                "id": 121442,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 6",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64620,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 6",
                "sectionName": "Group2Section 1",
                "groupName": "Group2",
                "groupId": 25492,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1
            },
            {
                "parentQuestionId": None,
                "id": 121444,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 1,
                "htmlString": "MS Question Randomization Low 7",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64620,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization Low 7",
                "sectionName": "Group2Section 1",
                "groupName": "Group2",
                "groupId": 25492,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1
            },
            {
                "parentQuestionId": None,
                "id": 121462,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high6",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64620,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high6",
                "sectionName": "Group2Section 1",
                "groupName": "Group2",
                "groupId": 25492,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1
            },
            {
                "parentQuestionId": None,
                "id": 121464,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 3,
                "htmlString": "MS Question Randomization high7",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64620,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization high7",
                "sectionName": "Group2Section 1",
                "groupName": "Group2",
                "groupId": 25492,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1
            },
            {
                "parentQuestionId": None,
                "id": 121484,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium6",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64620,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium6",
                "sectionName": "Group2Section 1",
                "groupName": "Group2",
                "groupId": 25492,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 1,
                "candidateAnswer": "B",
                "obtainedMark": -1
            },
            {
                "parentQuestionId": None,
                "id": 121486,
                "noOfAttachments": 0,
                "questionType": 7,
                "difficultyLevel": 4,
                "htmlString": "MS Question Randomization medium7",
                "config": None,
                "codingQuestionSubType": None,
                "sectionId": 64620,
                "typeOfQuestionText": "MCQ",
                "questionString": "MS Question Randomization medium7",
                "sectionName": "Group2Section 1",
                "groupName": "Group2",
                "groupId": 25492,
                "mark": 1,
                "correctAnswer": "A",
                "timeSpent": 17,
                "candidateAnswer": "B",
                "obtainedMark": -1
            }
        ],
        "questionTypeWiseOverall": {
            "mcq": {
                "rank": 5,
                "marks": -18,
                "averageMarks": 10.6,
                "highestMarks": 18,
                "lowestMarks": -18,
                "totalMarks": 18,
                "percentage": -100
            }
        },
        "questionWiseCodingSummary": {},
        "groupAndSectionWiseMarks": [
            {
                "id": 25490,
                "name": "Group1",
                "totalMarks": 12,
                "obtainedMarks": -12,
                "sectionInfo": [
                    {
                        "sectionId": 64616,
                        "sectionName": "Group1Section 1",
                        "totalMarks": 4.5,
                        "questionIds": [
                            121400,
                            121402,
                            121404,
                            121410,
                            121412,
                            121414,
                            121420,
                            121422,
                            121424
                        ],
                        "questionType": "MCQ",
                        "obtainedMarks": -4.5,
                        "difficultWiseCount": {
                            "low": {
                                "correct": None,
                                "inCorrect": 3,
                                "skipped": None,
                                "total": 3,
                                "timeSpent": 10,
                                "timeSpentCount": 3,
                                "partialCorrect": None
                            },
                            "medium": {
                                "correct": None,
                                "inCorrect": 3,
                                "skipped": None,
                                "total": 3,
                                "timeSpent": 3,
                                "timeSpentCount": 3,
                                "partialCorrect": None
                            },
                            "high": {
                                "correct": None,
                                "inCorrect": 3,
                                "skipped": None,
                                "total": 3,
                                "timeSpent": 3,
                                "timeSpentCount": 3,
                                "partialCorrect": None
                            }
                        },
                        "difficultyWiseMarks": {
                            "low": {
                                "maxMarks": 0.6000000000000001,
                                "obtainedMarks": -0.6000000000000001
                            },
                            "medium": {
                                "maxMarks": 1.5,
                                "obtainedMarks": -1.5
                            },
                            "high": {
                                "maxMarks": 2.4,
                                "obtainedMarks": -2.4
                            }
                        },
                        "questionCount": {
                            "total": 9,
                            "correct": None,
                            "inCorrect": 9,
                            "skipped": None,
                            "partialCorrect": None
                        }
                    },
                    {
                        "sectionId": 64618,
                        "sectionName": "Group1Section2",
                        "totalMarks": 7.5,
                        "questionIds": [
                            121406,
                            121408,
                            121416,
                            121418,
                            121426,
                            121428
                        ],
                        "questionType": "MCQ",
                        "obtainedMarks": -7.5,
                        "difficultWiseCount": {
                            "low": {
                                "correct": None,
                                "inCorrect": 2,
                                "skipped": None,
                                "total": 2,
                                "timeSpent": 2,
                                "timeSpentCount": 2,
                                "partialCorrect": None
                            },
                            "medium": {
                                "correct": None,
                                "inCorrect": 2,
                                "skipped": None,
                                "total": 2,
                                "timeSpent": 2,
                                "timeSpentCount": 2,
                                "partialCorrect": None
                            },
                            "high": {
                                "correct": None,
                                "inCorrect": 2,
                                "skipped": None,
                                "total": 2,
                                "timeSpent": 2,
                                "timeSpentCount": 2,
                                "partialCorrect": None
                            }
                        },
                        "difficultyWiseMarks": {
                            "low": {
                                "maxMarks": 2.1,
                                "obtainedMarks": -2.1
                            },
                            "medium": {
                                "maxMarks": 2.5,
                                "obtainedMarks": -2.5
                            },
                            "high": {
                                "maxMarks": 2.9,
                                "obtainedMarks": -2.9
                            }
                        },
                        "questionCount": {
                            "total": 6,
                            "correct": None,
                            "inCorrect": 6,
                            "skipped": None,
                            "partialCorrect": None
                        }
                    }
                ],
                "difficultWiseCount": {
                    "low": {
                        "correct": None,
                        "inCorrect": 5,
                        "skipped": None,
                        "total": 5,
                        "timeSpent": 12,
                        "timeSpentCount": 5,
                        "partialCorrect": None
                    },
                    "medium": {
                        "correct": None,
                        "inCorrect": 5,
                        "skipped": None,
                        "total": 5,
                        "timeSpent": 5,
                        "timeSpentCount": 5,
                        "partialCorrect": None
                    },
                    "high": {
                        "correct": None,
                        "inCorrect": 5,
                        "skipped": None,
                        "total": 5,
                        "timeSpent": 5,
                        "timeSpentCount": 5,
                        "partialCorrect": None
                    }
                },
                "difficultyWiseMarks": {
                    "low": {
                        "maxMarks": 2.7,
                        "obtainedMarks": -2.7
                    },
                    "medium": {
                        "maxMarks": 4,
                        "obtainedMarks": -4
                    },
                    "high": {
                        "maxMarks": 5.3,
                        "obtainedMarks": -5.3
                    }
                },
                "questionCount": {
                    "total": 15,
                    "correct": None,
                    "inCorrect": 15,
                    "skipped": None,
                    "partialCorrect": None
                }
            },
            {
                "id": 25492,
                "name": "Group2",
                "totalMarks": 6,
                "obtainedMarks": -6,
                "sectionInfo": [
                    {
                        "sectionId": 64620,
                        "sectionName": "Group2Section 1",
                        "totalMarks": 6,
                        "questionIds": [
                            121442,
                            121444,
                            121462,
                            121464,
                            121484,
                            121486
                        ],
                        "questionType": "MCQ",
                        "obtainedMarks": -6,
                        "difficultWiseCount": {
                            "low": {
                                "correct": None,
                                "inCorrect": 2,
                                "skipped": None,
                                "total": 2,
                                "timeSpent": 2,
                                "timeSpentCount": 2,
                                "partialCorrect": None
                            },
                            "medium": {
                                "correct": None,
                                "inCorrect": 2,
                                "skipped": None,
                                "total": 2,
                                "timeSpent": 18,
                                "timeSpentCount": 2,
                                "partialCorrect": None
                            },
                            "high": {
                                "correct": None,
                                "inCorrect": 2,
                                "skipped": None,
                                "total": 2,
                                "timeSpent": 2,
                                "timeSpentCount": 2,
                                "partialCorrect": None
                            }
                        },
                        "difficultyWiseMarks": {
                            "low": {
                                "maxMarks": 2,
                                "obtainedMarks": -2
                            },
                            "medium": {
                                "maxMarks": 2,
                                "obtainedMarks": -2
                            },
                            "high": {
                                "maxMarks": 2,
                                "obtainedMarks": -2
                            }
                        },
                        "questionCount": {
                            "total": 6,
                            "correct": None,
                            "inCorrect": 6,
                            "skipped": None,
                            "partialCorrect": None
                        }
                    }
                ],
                "difficultWiseCount": {
                    "low": {
                        "correct": None,
                        "inCorrect": 2,
                        "skipped": None,
                        "total": 2,
                        "timeSpent": 2,
                        "timeSpentCount": 2,
                        "partialCorrect": None
                    },
                    "medium": {
                        "correct": None,
                        "inCorrect": 2,
                        "skipped": None,
                        "total": 2,
                        "timeSpent": 18,
                        "timeSpentCount": 2,
                        "partialCorrect": None
                    },
                    "high": {
                        "correct": None,
                        "inCorrect": 2,
                        "skipped": None,
                        "total": 2,
                        "timeSpent": 2,
                        "timeSpentCount": 2,
                        "partialCorrect": None
                    }
                },
                "difficultyWiseMarks": {
                    "low": {
                        "maxMarks": 2,
                        "obtainedMarks": -2
                    },
                    "medium": {
                        "maxMarks": 2,
                        "obtainedMarks": -2
                    },
                    "high": {
                        "maxMarks": 2,
                        "obtainedMarks": -2
                    }
                },
                "questionCount": {
                    "total": 6,
                    "correct": None,
                    "inCorrect": 6,
                    "skipped": None,
                    "partialCorrect": None
                }
            }
        ],
        "assessment": {
            "id": 2452480,
            "percentage": -100,
            "testId": 15370,
            "testName": "Ms MCQ Marking Schema test",
            "totalMarks": 18,
            "attendedOn": "2022-09-05T16:59:43",
            "marksObtained": -18,
            "questionPaperId": 10780,
            "questionPaperTenantId": 159,
            "correct": 1,
            "isSuspicious": None,
            "reviewerComment": None,
            "timeSpent": 44,
            "testUserStatus": "Attended",
            "evalStatus": "Evaluated",
            "evalBy": "System",
            "evalOn": "2022-09-05T16:59:44",
            "isOffline": False,
            "typeOfTestId": 0,
            "isLoginCredentialSent": False,
            "conferenceId": None,
            "candidateId": 1412954,
            "typeOfTest": "Default",
            "highestScore": 18,
            "lowestScore": -18,
            "averageScore": 10.6,
            "scoreSummary": {
                "objective": {
                    "correct": 0,
                    "inCorrect": 21,
                    "skipped": 0,
                    "totalMarks": 18
                },
                "subjective": {
                    "obtained": 0,
                    "skipped": 0,
                    "totalMarks": 0
                },
                "coding": {
                    "obtained": 0,
                    "skipped": 0,
                    "totalMarks": 0
                }
            },
            "totalTestUsers": 21,
            "testUsersWithScore": 5,
            "candidateRank": 5,
            "percentile": 16.67
        },
        "tenantInfo": {
            "logoUrl": "https://s3.ap-southeast-1.amazonaws.com/test-all-hirepro-files/AT/applicationLogo/d24e14cc-c686-4358-9547-72f5314da123hirepro_logo_dark_1x.png?AWSAccessKeyId=AKIAIPKT65FIXSXQGDOQ&Signature=HvYBalPMR6nE9knhi3WvbuPFZUs%3D&Expires=1662446658"
        },
        "loginInfo": [
            {
                "messageLog": {},
                "clientSystemInfo": "Browser:chrome-104.0.0.0, OS:windows-windows-10, IPAddress:2401:4900:2344:e244:2584:b945:7ad4:cfac, timeZone:Asia/Calcutta,publicIp:106.208.136.144",
                "createdOn": "2022-09-05T16:58:51",
                "activityLog": {
                    "loginToken": "Tkn:73d7cf7e-fd55-41dd-85e0-ee51b2d3d273",
                    "candidateLoginActivityLog": [
                        {
                            "focusInOutLogInfo": {
                                "window": "parent",
                                "focusStat": "focusOut",
                                "timeStamp": "2022-09-05T11:28:59.290Z"
                            }
                        },
                        {
                            "focusInOutLogInfo": {
                                "window": "parent",
                                "focusStat": "focusIn",
                                "timeStamp": "2022-09-05T11:29:04.071Z"
                            }
                        },
                        {
                            "focusInOutLogInfo": {
                                "window": "parent",
                                "focusStat": "focusOut",
                                "timeStamp": "2022-09-05T11:29:35.674Z"
                            }
                        },
                        {
                            "focusInOutLogInfo": {
                                "window": "parent",
                                "focusStat": "focusIn",
                                "timeStamp": "2022-09-05T11:29:36.643Z"
                            }
                        },
                        {
                            "focusInOutLogInfo": {
                                "window": "parent",
                                "focusStat": "focusOut",
                                "timeStamp": "2022-09-05T11:29:37.756Z"
                            }
                        },
                        {
                            "focusInOutLogInfo": {
                                "window": "parent",
                                "focusStat": "focusIn",
                                "timeStamp": "2022-09-05T11:29:39.207Z"
                            }
                        },
                        {
                            "beforeSubmitLog": [
                                {
                                    "url": "/py/assessment/htmltest/api/v1/company-info/",
                                    "count": 1,
                                    "currentCallTime": 1662377330513
                                },
                                {
                                    "url": "/py/assessment/htmltest/api/v2/login_to_test/",
                                    "count": 1,
                                    "currentCallTime": 1662377331173
                                },
                                {
                                    "url": "/s3_cached/hirepro-content/hirepro/assessment/globalConfig.json?v=805016",
                                    "count": 1,
                                    "currentCallTime": 1662377331255
                                },
                                {
                                    "url": "/py/assessment/htmltest/api/v1/get-test-basic-info/",
                                    "count": 1,
                                    "currentCallTime": 1662377335470
                                },
                                {
                                    "url": "/py/assessment/htmltest/api/v1/initiate-tua/",
                                    "count": 1,
                                    "currentCallTime": 1662377338352
                                },
                                {
                                    "url": "/py/assessment/htmltest/api/v1/loadtest/",
                                    "count": 1,
                                    "currentCallTime": 1662377338553
                                },
                                {
                                    "url": "/py/assessment/htmltest/api/v1/setloginstats/",
                                    "count": 2,
                                    "currentCallTime": 1662377345508
                                }
                            ]
                        },
                        {
                            "mediaDeviceInfoLog": {
                                "mediaDevices": [
                                    {
                                        "audioinput": "Default - Microphone Array (Intel® Smart Sound Technology (Intel® SST))"
                                    },
                                    {
                                        "audioinput": "Communications - Microphone Array (Intel® Smart Sound Technology (Intel® SST))"
                                    },
                                    {
                                        "audioinput": "Microphone Array (Intel® Smart Sound Technology (Intel® SST))"
                                    },
                                    {
                                        "videoinput": "Integrated Webcam (1bcf:2b98)"
                                    },
                                    {
                                        "videoinput": "OBS Virtual Camera"
                                    },
                                    {
                                        "audiooutput": "Default - Speakers (Realtek(R) Audio)"
                                    },
                                    {
                                        "audiooutput": "Communications - Speakers (Realtek(R) Audio)"
                                    },
                                    {
                                        "audiooutput": "Speakers (Realtek(R) Audio)"
                                    }
                                ]
                            }
                        }
                    ]
                },
                "loginTime": "2022-09-05T16:58:51",
                "id": 3099772
            }
        ],
        "focusedOutDetails": [
            {
                "focusedOutTime": 4.781
            },
            {
                "focusedOutTime": 0.969
            },
            {
                "focusedOutTime": 1.451
            }
        ],
        "proctoring_details": {
            "proctoring_json": {
                "faceDetails": {},
                "behaviouralDetails": {
                    "comment": "No behavioural suspicious data found"
                }
            },
            "img_suspicious": None,
            "face_suspicious": False,
            "manual_suspicious": None,
            "reviewer_comment": None,
            "system_overall_decision": 0,
            "final_decision": False
        },
        "vendorDetails": None,
        "testUserScoresTypeWise": {
            "mcqAndRtcSectionTotalGroupByScore": {
                "18.0": 3,
                "17.0": 1,
                "-18.0": 1
            },
            "codingSectionTotalGroupByScore": {},
            "overallSectionTotalGroupByScore": {
                "18.0": 3,
                "17.0": 1,
                "-18.0": 1
            },
            "overallSectionTotalCount": 5
        },
        "groupSectionScoreSummary": [
            {
                "groupId": 25490,
                "maxMarks": 12,
                "minMarks": -12,
                "avgMarks": 7.2
            },
            {
                "groupId": 25492,
                "maxMarks": 6,
                "minMarks": -6,
                "avgMarks": 3.4
            },
            {
                "sectionId": 64616,
                "maxMarks": 4.5,
                "minMarks": -4.5,
                "avgMarks": 2.7
            },
            {
                "sectionId": 64618,
                "maxMarks": 7.5,
                "minMarks": -7.5,
                "avgMarks": 4.5
            },
            {
                "sectionId": 64620,
                "maxMarks": 6,
                "minMarks": -6,
                "avgMarks": 3.4
            }
        ]
    }
}
# total_mark = mark_details['data']['assessment']['totalMarks']
# total_obtained_mark = mark_details['data']['assessment']['marksObtained']
# group1_total_mark = mark_details['data']['groupAndSectionWiseMarks'][0]['totalMarks']
# # print(type(mark_details['data']['groupAndSectionWiseMarks'][0]['totalMarks']))
# group1_obtained_mark = mark_details['data']['groupAndSectionWiseMarks'][0]['obtainedMarks']
# group1_section1_obtained_mark = mark_details['data']['groupAndSectionWiseMarks'][0]['sectionInfo'][0]['obtainedMarks']
# group1_section1_total_mark = mark_details['data']['groupAndSectionWiseMarks'][0]['sectionInfo'][0]['totalMarks']
# group1_section2_obtained_mark = mark_details['data']['groupAndSectionWiseMarks'][0]['sectionInfo'][1]['obtainedMarks']
# group1_section2_total_mark = mark_details['data']['groupAndSectionWiseMarks'][0]['sectionInfo'][1]['totalMarks']
# group2_section1_obtained_mark = mark_details['data']['groupAndSectionWiseMarks'][1]['sectionInfo'][0]['obtainedMarks']
# group2_section1_total_mark = mark_details['data']['groupAndSectionWiseMarks'][1]['sectionInfo'][0]['totalMarks']
#
# print(total_mark)
# print(total_obtained_mark)
# print(group1_total_mark)
# print(group1_obtained_mark)
# print(group1_section1_obtained_mark)
# print(group1_section1_total_mark)
# print(group1_section2_obtained_mark)
# print(group1_section2_total_mark)
# print(group2_section1_obtained_mark)
# print(group2_section1_total_mark)
mark = {'parentQuestionId': None, 'id': 121400, 'noOfAttachments': 0, 'questionType': 7, 'difficultyLevel': 1,
        'htmlString': 'MS Question Randomization Low 1', 'config': None, 'codingQuestionSubType': None,
        'sectionId': 64616, 'typeOfQuestionText': 'MCQ', 'questionString': 'MS Question Randomization Low 1',
        'sectionName': 'Group1Section 1', 'groupName': 'Group1', 'groupId': 25490, 'mark': 0.1, 'correctAnswer': 'A',
        'timeSpent': 7, 'candidateAnswer': 'B', 'obtainedMark': -0.1}

all_questions = mark_details['data']['mcq']
id  = 121404
for question in all_questions:
    if id == question.get('id'):
        question_wise_obtained_mark = question['obtainedMark']

