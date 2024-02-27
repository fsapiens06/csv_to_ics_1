"""
This is a program that turns csv itineraries into ics file.
"""

import pandas as pd
import re

# translate course name to chinese or abbreviations
def course_abbr (eng):
    table = {
        "Introduction to Surgery" : "外概",
        "Rehabilitation Medicine" : "復健",
        "Radiotherapy" : "放射",
        "Nuclear medicine" : "核醫",
        "Clinical Skill" : "Intro",
        "Imaging Diagnosis" : "影診",
        "Otolarynglolgy" : "ENT",
        "Cardiovascular medicine" : "CV",
        "Pulmonary medicine" : "呼吸",
        "Endocrinology" : "Meta",
        "Gastroenterology" : "GI",
        "Neurology" : "神經",
        "Psychiatry" : "精神",
        "Dermatology" : "皮膚",
        "Orthopedics" : "骨科",
        "Allergy, Immunology, & Rheumatology" : "AIR",
        "Infectious Diseases" : "感染",
        "Nephrology" : "腎臟",
        "Urology" : "泌尿",
        "Hematology" : "血液",
        "Medical Oncology" : "腫瘤",
        "Pathology&Laboratory" : "病理",
        "Pharmacology" : "藥理",
        "基礎臨床技能訓練課程" : "BCS"}

    try:
        return table[eng]
    except:
        return eng

# set classroom
def get_classroom (course):
    table = {
        "lecture" : "綜二教室",
        "exam" : "綜二教室/綜三教室",
        "pbl" : "見分組清單",
        "bcs_location" : "見分組清單",
        "藥理學實驗" : "知行樓503教室",
        "整合醫學暨中醫學現代進展" : "教學大樓303教室"}

    try:
        return table[course]
    except:
        return ""

# set directory
input_dir = "112-2醫四行事曆-1130221.csv"
output_dir = "112-2醫四行事曆-1130221_轉檔.ics"

# read in csv file
raw_df = pd.read_csv(input_dir, parse_dates = True, encoding = "cp950")

# create ics framework
ics_df = pd.DataFrame(columns = ["Subject", "Start Date", "Start Time", "End Date", "End Time", "All Day", "Description", "Location"])

# define specific patterns in the subject column
# Standard format: [<Course>]:<Lec_code>, <Topic>,<Lecturer>
std_pattern = r'\[([^:]+)\]:([^,]+), (.+),([^,]+)'
# PBL format: PBL-<Topic>-<Round>
pbl_pattern = r'PBL-([^-]+)-([^-]+)'
# optional format: [選修]<Course_name>
optn_pattern = r'\[選修\]([^\]]+)'

# define description field format
# Standard format: [<Course>] <Topic> (<Lecturer>)
std_subject = "[{}] {} ({})"
std_no_lecr_subject = "[{}] {}"
# Exam format: [考試] <Exam_name>
exam_subject = "[!考試!] {}"
# PBL format: [PBL] <Topic_name> - <Round>
pbl_subject = "[PBL] {} - {}"
# Optional courses format: [<Course>] (選修)
optn_subject = "[{}] (選修)"

# define description field format
std_description = "科目：{}\n課程編號：{}\n主題：{}\n講師：{}\n"
exam_description = "科目：{}\n課程編號：{}\n名稱：{}\n負責人：{}\n"
pbl_description = "教案編號：{}\n回合：{}\n"

# iterate through raw data to generate ics file
ics_index = -1          # ics_index++ before 1st loop
last_subject = "--initialized--"       # for merging consecutive events

for raw_index, raw_row in raw_df.iterrows():
    raw_subject = raw_row["主旨"]

    # auto merge consecutive events
    if last_subject == raw_subject:
        ics_row = ics_df.iloc[ics_index]
        ics_row["End Date"] = raw_row["結束日期"]
        ics_row["End Time"] = raw_row["結束時間"]
        continue
    # exceptions for exams: different lec_code for same exam
    if last_subject in raw_subject and re.match(std_pattern, raw_subject):
        ics_row = ics_df.iloc[ics_index]
        ics_row["End Date"] = raw_row["結束日期"]
        ics_row["End Time"] = raw_row["結束時間"]
        continue
    else:
        # generate 1 new line in ics file
        ics_index += 1
        ics_df.loc[ics_index] = ""
        ics_row = ics_df.iloc[ics_index]

    # fill in basic info
    ics_row["Start Date"] = raw_row["開始日期"]
    ics_row["Start Time"] = raw_row["開始時間"]
    ics_row["End Date"] = raw_row["結束日期"]
    ics_row["End Time"] = raw_row["結束時間"]
    ics_row["All Day"] = raw_row["全天"]
    
    # separate infos in subject field 
    # use re.match to find the matches in the input string
    std_match = re.match(std_pattern, raw_subject)
    pbl_match = re.match(pbl_pattern, raw_subject)
    optn_match = re.match(optn_pattern, raw_subject)
 
    # regular courses and exams
    if std_match:
        # extract values from matched groups
        course = std_match.group(1)
        lec_code = std_match.group(2)
        topic = std_match.group(3)
        lecturer = std_match.group(4)

        # courses without lecturer
        if lecturer == "-":
            ics_row["Subject"] = std_no_lecr_subject.format(course_abbr(course), topic)
            ics_row["Location"] = get_classroom("unknown")
            ics_row["Description"] = std_description.format(course, lec_code, topic, lecturer)
            
        # exams
        elif "考" in topic:
            # merging mode ON
            last_subject = topic
            
            ics_row["Subject"] = exam_subject.format(topic)
            ics_row["Location"] = get_classroom("exam")
            ics_row["Description"] = exam_description.format(course, lec_code, topic, lecturer)

        # courses with standard format
        else:
            ics_row["Subject"] = std_subject.format(course_abbr(course), topic, lecturer)
            ics_row["Location"] = get_classroom("lecture")
            ics_row["Description"] = std_description.format(course, lec_code, topic, lecturer)

    # PBL
    elif pbl_match:
        # merging mode ON
        last_subject = raw_subject

        # extract values from matched groups
        topic_name = pbl_match.group(1)
        n_round = pbl_match.group(2)
        
        ics_row["Subject"] = pbl_subject.format(topic_name, n_round)
        ics_row["Location"] = get_classroom("pbl")
        ics_row["Description"] = pbl_description.format(topic_name, n_round)
   
    elif raw_subject == "一對一訪談":
        ics_row["Subject"] = "[PBL] 一對一訪談"
        ics_row["Location"] = get_classroom("pbl")

    # Optional courses
    elif optn_match:
        course_name = optn_match.group(1)
        ics_row["Subject"] = optn_subject.format(course_name)
        ics_row["Location"] = get_classroom(course_name)

    # BCS
    elif raw_subject == "基礎臨床技能訓練課程":
        # merging mode ON
        last_subject = raw_subject 
        ics_row["Subject"] = "[BCS]"
        ics_row["Location"] = get_classroom("bcs")
    
    # other events
    else:
        # merging mode ON
        last_subject = raw_subject
        ics_row["Subject"] = raw_subject

# output ics file
ics_df.to_csv(output_dir, index = False, encoding = "utf-8")

