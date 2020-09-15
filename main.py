import json
import requests
import openpyxl
import era
from datetime import datetime

# スプレッドシートを開く
spreadsheet = openpyxl.load_workbook('./data/20200915data.xlsx')

print(spreadsheet.sheetnames)

# シートの内容を代入
patients_sheet = spreadsheet['陽性者の属性']
pcr_sheet = spreadsheet['PCR検査件数']
news_sheet = spreadsheet['最新の情報']

# 変数を初期化
patients_data = {
  "data": [],
  "last_update": "2020/09/15 20:00"
}

patients_summary = {
  "data": [],
  "last_update": "2020/09/15 20:00"
}

pcr_data = {}
news_data = {}

last_update = datetime.strptime('20200915' + '1900', '%Y%m%d%H%M')

print(last_update)

print(patients_sheet.max_row)

for i in range(patients_sheet.max_row) :
  date = era.reiwa_to_datetime(patients_sheet.cell(row=patients_sheet.max_row - i, column=2).value)
  date_shorted = str(date.year).zfill(4) + "-" + str(date.month).zfill(2) + "-" + str(date.day).zfill(2)

  number_data = str(patients_sheet.cell(row=patients_sheet.max_row - i, column=1).value).replace("例目", "")
  number = int(number_data) if (number_data == "-") == False else None

  address = patients_sheet.cell(row=patients_sheet.max_row - i, column=5).value

  age_data = str(patients_sheet.cell(row=patients_sheet.max_row - i, column=3).value)
  if age_data == "-" :
    age = ""
  elif "以上" in age_data :
    age = age_data.replace("以上", "歳以上")
  elif age_data == "園児" :
    age = age_data
  else :
    age = age_data + "代"

  gender_data = patients_sheet.cell(row=patients_sheet.max_row - i, column=4).value
  if gender_data == "-" :
    gender = ""
  else :
    gender = gender_data

  left_hospital_data = patients_sheet.cell(row=patients_sheet.max_row - i, column=6).value
  left_hospital = left_hospital_data if left_hospital_data != None else None

  patients_data["data"].append(
    {
      "No": number,
      "リリース日": date_shorted + "T08:00:00.000Z",
      "居住地": address,
      "年代と性別": age + gender,
      "退院": left_hospital,
      "date": date_shorted,
    }
  )

print(patients_data)

with open('./outputs/patients.json', 'w') as f:
  json.dump(patients_data, f, indent=4, ensure_ascii=False)