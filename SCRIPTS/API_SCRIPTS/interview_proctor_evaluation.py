from SCRIPTS.CRPO_COMMON.crpo_common import *
from SCRIPTS.CRPO_COMMON.credentials import *
from SCRIPTS.COMMON.read_excel import *
from SCRIPTS.COMMON.write_excel_new import *
from SCRIPTS.COMMON.io_path import *


class ProctorEvaluation:
    def __init__(self):
        write_excel_object.save_result(output_interview_proctor_evaluation)
        # 0th Row Header
        header = ['Proctoring Evaluation automation']
        # 1 Row Header
        write_excel_object.write_headers_for_scripts(0, 0, header, write_excel_object.black_color_bold)
        header = ['TestCases', 'Status', 'Candidate ID', 'Job Id', 'IR ID', 'Interviewer ID',
                  'Exp – Candidate Image status', 'Act – Candidate Image Status', 'Exp – Candidate Video Status',
                  'Act – Candidate Video Status', 'Exp – Candidate Overall Status', 'Act – Candidate Overall status',
                  'Exp  - Candidate Overall Value', 'Act – Candidate Overall Value', 'Exp - Candidate Browser status', 'Act - Candidate Browser Status',
                  'Exp - Candidate Lipsync Status', 'Act - Candidate Lipsync Status', 'Exp – interviewer image status',
                  'Act – Interviewer Image status', 'Exp – Interviewer Overall Status',
                  'Act – Interviewer overall status', 'Exp – interviewer overall value',
                  'Act – Interviewer Overall Value']
        write_excel_object.write_headers_for_scripts(1, 0, header, write_excel_object.black_color_bold)

    def proctor_detail(self, current_excel_data, token):
        row_count = 1
        for interview_details in current_excel_data:
            row_count = row_count + 1
            run_proctoring = CrpoCommon.run_proctoring(token, interview_details.get('irId'))
            if run_proctoring.get('status') == 'OK':
                candidate = run_proctoring['data']['interviewee'][0]
                candidate_video_suspicious_status = candidate['faceSuspicious']
                if candidate_video_suspicious_status:
                    candidate_video_suspicious_status = 'Suspicious'
                else:
                    candidate_video_suspicious_status = 'Not Suspicious'

                candidate_face_match_status = candidate['imgSuspicious']
                if candidate_face_match_status:
                    candidate_face_match_status = 'Suspicious'
                else:
                    candidate_face_match_status = 'Not Suspicious'

                candidate_overall_decision = candidate['systemOverallDecision']
                # print(candidate)
                # print(candidate_video_suspicious_status)
                # print(candidate_face_match_status)
                # print(candidate_overall_decision)
                candidate_final_status = 'Not Calculated'

                if candidate_overall_decision is not None:
                    if candidate_overall_decision >= 0.7:
                        candidate_final_status = 'Highly Suspicious'
                    elif candidate_overall_decision > 0:
                        candidate_final_status = 'Medium'
                    else:
                        candidate_final_status = 'Not Suspicious'
                else:
                    candidate_final_status = 'NA'
                # print(candidate_final_status)

                interviewer = run_proctoring['data']['interviewer'][0]
                interviewer_face_match_status = interviewer['imgSuspicious']
                if interviewer_face_match_status:
                    interviewer_face_match_status = 'Suspicious'
                else:
                    interviewer_face_match_status = 'Not Suspicious'
                interviewer_overall_decision = interviewer['systemOverallDecision']
                interviewer_final_status = 'Not Calculated'

                if interviewer_overall_decision is not None:
                    if interviewer_overall_decision >= 0.7:
                        interviewer_final_status = 'Highly Suspicious'
                    elif interviewer_overall_decision > 0:
                        interviewer_final_status = 'Suspicious'
                    else:
                        interviewer_final_status = 'Not Suspicious'
                else:
                    interviewer_final_status = 'NA'
            else:
                print("Run Proctoring is failed")
                candidate_video_suspicious_status = 'Failed'
                candidate_face_match_status = 'Failed'
                candidate_overall_decision = None
                candidate_final_status = 'Failed'
                interviewer_face_match_status = 'Failed'
                interviewer_overall_decision = None
                interviewer_final_status = 'Failed'

            lip_sync = CrpoCommon.lip_sync(token, interview_details.get('irId'))
            print(lip_sync)
            lip_sync_status = lip_sync['data']['suspiciousResult']

            write_excel_object.current_status_color = write_excel_object.green_color
            write_excel_object.current_status = "Pass"
            write_excel_object.compare_results_and_write_vertically(interview_details.get('testCases'), None, row_count,
                                                                    0)
            # write_excel_object.compare_results_and_write_vertically(interview_details.get('testCases'), None, row_count,
            #                                                         1)
            write_excel_object.compare_results_and_write_vertically(interview_details.get('candidateId'), None,
                                                                    row_count, 2)
            write_excel_object.compare_results_and_write_vertically(interview_details.get('jobId'), None, row_count, 3)
            write_excel_object.compare_results_and_write_vertically(interview_details.get('irId'), None, row_count, 4)
            write_excel_object.compare_results_and_write_vertically(interview_details.get('interviewerId'), None,
                                                                    row_count, 5)

            write_excel_object.compare_results_and_write_vertically(interview_details.get('candidateImageStatus'),
                                                                    candidate_face_match_status, row_count, 6)

            write_excel_object.compare_results_and_write_vertically(interview_details.get('candidateVideoStatus'),
                                                                    candidate_video_suspicious_status, row_count, 8)

            write_excel_object.compare_results_and_write_vertically(interview_details.get('candidateOverAllStatus'),
                                                                    candidate_final_status, row_count, 10)

            write_excel_object.compare_results_and_write_vertically(
                interview_details.get('candidateOverallSuspiciousValue'), candidate_overall_decision, row_count, 12)

            write_excel_object.compare_results_and_write_vertically(
                interview_details.get('candidateBrowserStatus'), 'API is not Returning data - under progress',
                row_count, 14)

            write_excel_object.compare_results_and_write_vertically(
                interview_details.get('candidateLipsyncStatus'), lip_sync_status, row_count, 16)

            write_excel_object.compare_results_and_write_vertically(
                interview_details.get('interviewerImageStatus'), interviewer_face_match_status, row_count, 18)

            write_excel_object.compare_results_and_write_vertically(
                interview_details.get('interviewerOverallSuspiciousStatus'), interviewer_final_status, row_count, 20)

            write_excel_object.compare_results_and_write_vertically(
                interview_details.get('interviewerOverallSuspiciousValue'), interviewer_overall_decision, row_count, 22)



            write_excel_object.compare_results_and_write_vertically(write_excel_object.current_status, None, row_count,
                                                                    1)

        write_excel_object.write_overall_status(testcases_count=1)


login_token = crpo_common_obj.login_to_crpo(cred_crpo_admin_crpodemo.get('user'),
                                            cred_crpo_admin_crpodemo.get('password'),
                                            cred_crpo_admin_crpodemo.get('tenant'))
excel_read_obj.excel_read(input_interview_proctoring_evaluation, 0)
excel_data = excel_read_obj.details
over_all_status = 'Pass'
proctor_obj = ProctorEvaluation()
# tuids = []

proctor_obj.proctor_detail(excel_data, login_token)
