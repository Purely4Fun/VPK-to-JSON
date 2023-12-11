import pandas as pd
import json
from datetime import datetime
import re

current_datetime = datetime.now()
current_month = current_datetime.month
current_year = current_datetime.year

if current_month > 8 and current_month <= 12 or current_month == 1:
    semester = 1
else:
    semester = 2

year = str(current_year)
past_year = str(current_year - 1)
future_year = str(current_year + 1)

if semester == 1:
    ext_year = year + future_year[-2:]
    future_year = future_year[-2:]
    year = year[-2:]
    two_years = year + "/" + future_year
    grade_year = year
else:
    ext_year = past_year + year[-2:]
    past_year = past_year[-2:]
    year = year[-2:]
    two_years = past_year + "/" + year
    grade_year = past_year

file_path = 'D:\JsonProject\Копия ВПК ИКТИБ.xlsx'
sheet_names = ['Teachers', 'Bachelor', 'Class', 'AllInfo', 'TSnils']

dfs = [pd.read_excel(file_path, sheet_name=sheet_name, header=None) for sheet_name in sheet_names]

teachers = dfs[0].values
groups = dfs[1].values

classes = dfs[2].ffill().dropna().values
students = dfs[3].values
snils = dfs[4].values

teachers_dict = {}
for row in classes:
        subject = row[1]
        teacher = row[5]
        if subject in teachers_dict:
            teachers_dict[subject].append(teacher)
        else:
            teachers_dict[subject] = [teacher]

arrGr = []
count = 1
for i in classes:
    
    snilsArr=[]
    studArr = []
    teacherArr = []
    vpk_name = i[0].split('(')[0].strip()    
    subject = i[1]
    teacher = i[5]
    for key, value in teachers_dict.items():
        if subject == key and teacher in value:
            for k in snils:
                for o in value:
                    if k[0] == o:
                        if not any(d.get("id") == k[2] for d in snilsArr):
                            snilsArr.append({"id": k[2]})
    vpkTmp = "VPK" + re.findall(r'\d+', vpk_name)[0]
    globalID = f"{ext_year}-{semester}_{int(i[3])}-{int(i[4])}_{vpkTmp}_LP_S1"
    if not any(d.get("globalExternalID") == globalID for d in arrGr):
        globalID = f"{ext_year}-{semester}_{int(i[3])}-{int(i[4])}_{vpkTmp}_LP_S1"
    else:
        globalID = f"{ext_year}-{semester}_{int(i[3])}-{int(i[4])}_{vpkTmp}_{count}_LP_S1"
        count+=1
    grade = f"{vpk_name}-{grade_year}-{semester}_{i[1]}_ЛП_о/1"
    vpk_name = f"{vpk_name}-о-{two_years}-{semester} {i[1]} Лекции+Практика"
    found = False
    for j in groups:
      if j[6] in i[0]:
        tempStudList = dfs[3][dfs[3][1]==j[0]][13].tolist()
        for n in studArr:
            if tempStudList and tempStudList[0] == n.get("stu1c_id"):
                found = True
                break
        if not found and tempStudList:
          tempStud = {"stu1c_id": dfs[3][dfs[3][1]==j[0]][13].tolist()[0], "id": dfs[3][dfs[3][1]==j[0]][11].tolist()[0],
                        "subdivisionid": dfs[3][dfs[3][1]==j[0]][25].tolist()[0], "planid": dfs[3][dfs[3][1]==j[0]][26].tolist()[0]}
          studArr.append(tempStud)  
    found = False
    for j in arrGr:
        if vpk_name == j.get("teams_name"):
            found = True
            break
    if not found:
        if str(i[6])=="диф.зач.":
           t = "Дифференцированный зачет"
        elif str(i[6])=="зач.":
           t = "Зачет"
        else:
            t = "Экзамен"
        if not any(d.get("globalExternalID") == globalID for d in arrGr):
            tempGr = {"teams_name": vpk_name, "subject": i[1], "name": grade, "subgroup": "1", "faculty": "000000080",
                  "externalID": str(int(i[2])), "globalExternalID": globalID, "teachers": snilsArr, "type": t, "students": studArr}
            arrGr.append(tempGr)
        else:
            tempGr = {"teams_name": vpk_name, "subject": i[1], "name": grade, "subgroup": "1", "faculty": "000000080",
                  "externalID": str(int(i[2])), "globalExternalID": globalID, "teachers": snilsArr, "type": t, "students": studArr}
            arrGr.append(tempGr)
            
                              
data = {"year": str(current_year), "semester": str(semester), "groups": arrGr }
json_data = json.dumps(data, ensure_ascii=False)
with open('test.json', 'w', encoding='utf-8') as outfile:
    json.dump(data, outfile, ensure_ascii=False)
print(json_data)