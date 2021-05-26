import xlrd
import datetime


class Excel_Data:
    def __init__(self):
        self.xl_candidate_name = []
        self.xl_candidate_fname = []
        self.xl_candidate_mname = []
        self.xl_candidate_lname = []
        self.xl_email = []
        self.xl_mobile = []
        self.xl_phone_office = []
        self.xl_marital = []
        self.xl_gender = []
        self.xl_dob = []

        self.xl_pancard = []
        self.xl_passport = []
        self.xl_aadhar = []

        self.xl_usn = []
        self.xl_college = []
        self.xl_degree = []
        self.xl_current_Location = []
        self.xl_total_experience = []

        self.xl_linkedin = []
        self.xl_facebook = []
        self.xl_twitter = []

        self.xl_text1 = []
        self.xl_text2 = []
        self.xl_text3 = []
        self.xl_text4 = []
        self.xl_text5 = []
        self.xl_text6 = []
        self.xl_text7 = []
        self.xl_text8 = []
        self.xl_text9 = []
        self.xl_text10 = []
        self.xl_text11 = []
        self.xl_text12 = []
        self.xl_text13 = []
        self.xl_text14 = []
        self.xl_text15 = []

        self.xl_integer1 = []
        self.xl_integer2 = []
        self.xl_integer3 = []
        self.xl_integer4 = []
        self.xl_integer5 = []
        self.xl_integer6 = []
        self.xl_integer7 = []
        self.xl_integer8 = []
        self.xl_integer9 = []
        self.xl_integer10 = []
        self.xl_integer11 = []
        self.xl_integer12 = []
        self.xl_integer13 = []
        self.xl_integer14 = []
        self.xl_integer15 = []
        self.xl_duplication_rule_json = []
        self.xl_duplicatetext = []
        self.xl_expected = []
        self.xl_expected_message = []

    def Data_read(self):
        wb = xlrd.open_workbook('/home/muthumurugan/Desktop/Automation/PythonWorkingScripts_InputData/'
                                'CRPO-RPO/Duplication_Rule/Rpotest_New1.xls')
        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        # print(sheetname)
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range(1, sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            self.xl_candidate_name.append(str(rows[0]))
            self.xl_candidate_fname.append(str(rows[1]))
            self.xl_candidate_mname.append(str(rows[2]))
            self.xl_candidate_lname.append(str(rows[3]))

            self.xl_email.append(str(rows[4]))

            if not rows[5]:
                self.xl_mobile.append(str(rows[5]))
            else:
                self.xl_mobile.append(int(rows[5]))

            if not rows[6]:
                self.xl_phone_office.append(str(rows[6]))
            else:
                self.xl_phone_office.append(int(rows[6]))

            if not rows[7]:
                self.xl_marital.append(str(rows[7]))
            else:
                self.xl_marital.append(int(rows[7]))

            if not rows[8]:
                self.xl_gender.append(str(rows[8]))
            else:
                self.xl_gender.append(int(rows[8]))

            if not rows[9]:
                self.xl_dob.append(str(''))
            else:
                created_on_to = sh1.cell_value(rowx=(i), colx=9)
                self.created_on_to = datetime.datetime(*xlrd.xldate_as_tuple(created_on_to, wb.datemode))
                self.created_on_to = self.created_on_to.strftime("%Y-%m-%d")
                self.xl_dob.append(self.created_on_to)

            self.xl_pancard.append(str(rows[10]))
            self.xl_passport.append(str(rows[11]))

            if not rows[12]:
                self.xl_aadhar.append(str(rows[12]))
            else:
                self.xl_aadhar.append(int(rows[12]))

            self.xl_usn.append(str(rows[13]))

            if not rows[14]:
                self.xl_college.append(str(rows[14]))
            else:
                self.xl_college.append(int(rows[14]))

            if not rows[15]:
                self.xl_degree.append(str(rows[15]))
            else:
                self.xl_degree.append(int(rows[15]))

            if not rows[16]:
                self.xl_current_Location.append(str(rows[16]))
            else:
                self.xl_current_Location.append(int(rows[16]))

            if not rows[17]:
                self.xl_total_experience.append(str(rows[17]))
            else:
                self.xl_total_experience.append(int(rows[17]))

            self.xl_linkedin.append(str(rows[18]))
            self.xl_facebook.append(str(rows[19]))
            self.xl_twitter.append(str(rows[20]))

            self.xl_text1.append(str(rows[21]))
            self.xl_text2.append(str(rows[22]))
            self.xl_text3.append(str(rows[23]))
            self.xl_text4.append(str(rows[24]))
            self.xl_text5.append(str(rows[25]))
            self.xl_text6.append(str(rows[26]))
            self.xl_text7.append(str(rows[27]))
            self.xl_text8.append(str(rows[28]))
            self.xl_text9.append(str(rows[29]))
            self.xl_text10.append(str(rows[30]))
            self.xl_text11.append(str(rows[31]))
            self.xl_text12.append(str(rows[32]))
            self.xl_text13.append(str(rows[33]))
            self.xl_text14.append(str(rows[34]))
            self.xl_text15.append(str(rows[35]))

            if not rows[36]:
                self.xl_integer1.append(str(rows[36]))
            else:
                self.xl_integer1.append(int(rows[36]))

            if not rows[37]:
                self.xl_integer2.append(str(rows[37]))
            else:
                self.xl_integer2.append(int(rows[37]))

            if not rows[38]:
                self.xl_integer3.append(str(rows[38]))
            else:
                self.xl_integer3.append(int(rows[38]))

            if not rows[39]:
                self.xl_integer4.append(str(rows[39]))
            else:
                self.xl_integer4.append(int(rows[39]))

            if not rows[40]:
                self.xl_integer5.append(str(rows[40]))
            else:
                self.xl_integer5.append(int(rows[40]))

            if not rows[41]:
                self.xl_integer6.append(str(rows[41]))
            else:
                self.xl_integer6.append(int(rows[41]))

            if not rows[42]:
                self.xl_integer7.append(str(rows[42]))
            else:
                self.xl_integer7.append(int(rows[42]))

            if not rows[43]:
                self.xl_integer8.append(str(rows[43]))
            else:
                self.xl_integer8.append(int(rows[43]))

            if not rows[44]:
                self.xl_integer9.append(str(rows[44]))
            else:
                self.xl_integer9.append(int(rows[44]))

            if not rows[45]:
                self.xl_integer10.append(str(rows[45]))
            else:
                self.xl_integer10.append(int(rows[45]))

            if not rows[46]:
                self.xl_integer11.append(str(rows[46]))
            else:
                self.xl_integer11.append(int(rows[46]))

            if not rows[47]:
                self.xl_integer12.append(str(rows[47]))
            else:
                self.xl_integer12.append(int(rows[47]))

            if not rows[48]:
                self.xl_integer13.append(str(rows[48]))
            else:
                self.xl_integer13.append(int(rows[48]))

            if not rows[49]:
                self.xl_integer14.append(str(rows[49]))
            else:
                self.xl_integer14.append(int(rows[49]))

            if not rows[50]:
                self.xl_integer15.append(str(rows[50]))
            else:
                self.xl_integer15.append(int(rows[50]))

            self.xl_duplication_rule_json.append(str(rows[51]))
            self.xl_duplicatetext.append(str(rows[52]))
            self.xl_expected.append(str(rows[53]))
            self.xl_expected_message.append(rows[54])

    def disp(self):
        print  self.xl_candidate_name
        print self.xl_candidate_fname
        print self.xl_candidate_mname
        print self.xl_candidate_lname
        print self.xl_email
        print self.xl_mobile
        print self.xl_phone_office
        print  self.xl_marital
        print self.xl_gender
        print self.xl_dob
        print  self.xl_pancard
        print self.xl_passport
        print self.xl_aadhar
        print  self.xl_usn
        print self.xl_college
        print self.xl_degree
        print  self.xl_current_Location
        print self.xl_total_experience
        print self.xl_linkedin
        print  self.xl_facebook
        print self.xl_twitter

        print self.xl_text1
        print self.xl_text2
        print self.xl_text3
        print self.xl_text4
        print self.xl_text5
        print self.xl_text6
        print self.xl_text7
        print self.xl_text8
        print self.xl_text9
        print self.xl_text10
        print self.xl_text11
        print self.xl_text12
        print self.xl_text13
        print self.xl_text14
        print self.xl_text15

        print self.xl_integer1
        print self.xl_integer2
        print self.xl_integer3
        print self.xl_integer4
        print self.xl_integer5
        print self.xl_integer6
        print self.xl_integer7
        print self.xl_integer8
        print self.xl_integer9
        print self.xl_integer10
        print self.xl_integer11
        print self.xl_integer12
        print self.xl_integer13
        print self.xl_integer14
        print self.xl_integer15

        print self.xl_duplication_rule_json
        print self.xl_duplicatetext
        print self.xl_expected_message


dup_ob = Excel_Data()
dup_ob.Data_read()
# dup_ob.disp()
