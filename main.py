from datetime import datetime, timedelta
from io import BytesIO

import json
import openpyxl
import era
import gmailrestwrapper
import base64
import sys

client_id = (sys.argv)[1]
client_secret = (sys.argv)[2]
refresh_token = (sys.argv)[3]
kyoto_addr = (sys.argv)[4]
my_addr = (sys.argv)[5]

param = {
  "client_id": client_id,
  "client_secret": client_secret,
  "refresh_token": refresh_token,
  "grant_type": "refresh_token"
}

today = datetime.now()

token = gmailrestwrapper.get_token(param)
msg_list = gmailrestwrapper.get_message_list(
  token,
  query=(
    "from:" + 
    kyoto_addr + 
    " OR from:" + 
    my_addr + 
    " after:" +
    str(today.year).zfill(4) + "/" + str(today.month).zfill(2) + "/" + str(today.day).zfill(2) +
    " before:" +
    str(today.year).zfill(4) + "/" + str(today.month).zfill(2) + "/" + str(today.day + 1).zfill(2)
    ),
  maxResults=20)

last_mail_date = None

if ("messages" in msg_list.keys()) == False:
  exit()
else:
  for msg_id_in_list in msg_list["messages"]:
    msg_id = msg_id_in_list["id"]
    msg = gmailrestwrapper.get_message(token, msg_id)
    date = int(int(msg["internalDate"]) / 1000)
    date_datetime = datetime.utcfromtimestamp(date) + timedelta(hours=9)

  # メールから xlsx ファイルを取り出す
  for payload in msg["payload"]["parts"]:
    if (last_mail_date == None):
      last_mail_date = date_datetime
    elif date_datetime >= last_mail_date:
      last_mail_date = date_datetime

    if ("data.xlsx" in payload["filename"]) and (len(payload["filename"]) == 17) and (date_datetime >= last_mail_date):
      last_update = date_datetime
      last_update = (
        str(last_update.year).zfill(4) + "-" +
        str(last_update.month).zfill(2) + "-" +
        str(last_update.day).zfill(2) + "T" +
        str(last_update.hour).zfill(2) + ":" +
        str(last_update.minute).zfill(2) + ":" +
        str(last_update.second).zfill(2) +
        ".000Z"
      )
      filename = payload["filename"]
      spread_data = base64.urlsafe_b64decode((json.loads(gmailrestwrapper.get_attachment(token, msg_id, payload["body"]["attachmentId"])))["data"].encode("ascii"))



# exit()

# スプレッドシートを開く
spreadsheet = openpyxl.load_workbook(filename=BytesIO(spread_data))

# シートの内容を代入
patients_sheet = spreadsheet['陽性者の属性']
pcr_sheet = spreadsheet['PCR検査件数']
news_sheet = spreadsheet['最新の情報']

# 変数を初期化
patients_data = {
  "data": [],
  "last_update": last_update
}

patients_data_converted = {
  "data": [],
  "last_update": last_update
}

patients_summary = {
  "data": [],
  "last_update": last_update
}

inspections_summary = {
  "data": [],
  "last_update": last_update
}

pcr_data = {}
news_data = {}




patient_count = 0
patients_count_day = 0

recent_patient_date = None

zero_days_sum = 0


# 感染者のデータを生成
for i in range(patients_sheet.max_row) :
  date = (era.reiwa_to_datetime(patients_sheet.cell(row=patients_sheet.max_row - i, column=2).value)) + timedelta(hours=8)
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

# 患者の一覧を日付でソート
patients_sorted = sorted(patients_data["data"], key=lambda x: datetime.strptime(x["リリース日"], "%Y-%m-%dT%H:%M:%S.000Z"))

# ソートした患者の一覧をもとに日毎の患者の小計を計算
for i in range(len(patients_sorted)) :

  date = datetime.strptime(patients_sorted[i]["リリース日"], "%Y-%m-%dT%H:%M:%S.000Z")

  if i == 0 :
    recent_patient_date = date

  if (recent_patient_date != date) :
    patients_summary["data"].append(
      {
        "日付": 
          str(recent_patient_date.year).zfill(4) + "-" + 
          str(recent_patient_date.month).zfill(2) + "-" +
          str(recent_patient_date.day).zfill(2) + "T08:00:00.000Z",
        "小計": patient_count
      }
    )
    patient_count = 0
 
  patient_count += 1

  if i + 1 == patients_sheet.max_row :
    patients_summary["data"].append(
      {
        "日付": 
          str(date.year).zfill(4) + "-" + 
          str(date.month).zfill(2) + "-" +
          str(date.day).zfill(2) + "T08:00:00.000Z",
        "小計": patient_count
      }
    )

  zero_days = (date - recent_patient_date).days

  if zero_days >= 2 :
    for j in range(zero_days - 1) :
      zero_date = recent_patient_date + timedelta(days=j+1)
      patients_summary["data"].append(
        {
          "日付": str(zero_date.year).zfill(4) + "-" + str(zero_date.month).zfill(2) + "-" + str(zero_date.day).zfill(2) + "T08:00:00.000Z",
          "小計": 0
        }
      )

  recent_patient_date = date


# PCR 検査件数の小計を計算

main_summary = {
  "attr": "検査実施人数",
  "value": pcr_sheet.cell(row=1, column=2).value,
  "children": [
    {
      "attr": "陽性患者数",
      "value": pcr_sheet.cell(row=1, column=3).value,
      "children": [
        {
          "attr": "入院中・入院調整中",
          "value": pcr_sheet.cell(row=1, column=5).value
        },
        {
          "attr": "宿泊施設",
          "value": pcr_sheet.cell(row=1, column=7).value
        },
        {
          "attr": "自宅療養",
          "value": pcr_sheet.cell(row=1, column=8).value
        },
        {
          "attr": "死亡",
          "value": pcr_sheet.cell(row=1, column=9).value
        },
        {
          "attr": "退院・解除",
          "value": pcr_sheet.cell(row=1, column=4).value
        },
      ]
    }
  ],
  "last_update": last_update
}

last_inspected = 0

# PCR 検査実施数の日毎の集計
for i in range(pcr_sheet.max_row) :
  
  date = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=1).value + timedelta(hours=8)
  inspected = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=2).value

  if inspected != None :
    inspections_summary["data"].append(
      {
        "日付": 
          str(date.year).zfill(4) + "-" + 
          str(date.month).zfill(2) + "-" +
          str(date.day).zfill(2) + "T08:00:00.000Z",
        "小計": inspected - last_inspected 
      }
    )

  if inspected != None:
    last_inspected = inspected

del inspections_summary["data"][0]

for i in range(pcr_sheet.max_row) :
  date = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=1).value + timedelta(hours=8)
  inspected = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=2).value
  patients = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=3).value
  left_hospital = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=4).value
  in_hospital = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=5).value
  seriously_patients = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=6).value
  in_hotel = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=7).value
  in_home = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=8).value
  died = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=9).value
  adjusting = pcr_sheet.cell(row=pcr_sheet.max_row - i, column=10).value
  # print(date, inspected, patients, left_hospital, in_hospital, seriously_patients, in_hotel, in_home, died, adjusting)

#print(patients_sorted)

with open('patients.json', 'w') as f:
  json.dump(patients_data, f, indent=4, ensure_ascii=False)

with open('patients_summary.json', 'w') as f:
  json.dump(patients_summary, f, indent=4, ensure_ascii=False)

with open('last_update.json', 'w') as f:
  json.dump(
    {
      "last_update": last_update
    },
    f, indent=4, ensure_ascii=False
  )

with open('main_summary.json', 'w') as f:
  json.dump(main_summary, f, indent=4, ensure_ascii=False)

with open('inspections_summary.json', 'w') as f:
  json.dump(inspections_summary, f, indent=4, ensure_ascii=False)

