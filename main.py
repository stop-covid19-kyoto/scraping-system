from datetime import datetime, timedelta
from io import BytesIO

import json
import openpyxl
import era
import gmailrestwrapper
import base64
import sys

class data_generator:
  # datetime型のオブジェクトから短い日付を返す
  def get_shorted_date(self, datetime_obj):
    return (
      str(datetime_obj.year).zfill(4)
      + "-" + 
      str(datetime_obj.month).zfill(2)
      + "-" + 
      str(datetime_obj.day).zfill(2)
    )

  # 陽性者の一覧を生成
  def get_patients_data(self, sheet):
    patients_data = []
    for i in range(sheet.max_row) :

      date = (
        era.reiwa_to_datetime(
          sheet.cell(
            row=sheet.max_row - i,
            column=2
          ).value
        ) + timedelta(hours=8)
      )

      shorted_date = self.get_shorted_date(date)

      patient = str(
        sheet.cell(
          row=sheet.max_row - i,
          column=1
        ).value
      ).replace("例目", "")

      patient_number = (
        int(patient) if (patient == "-") == False else None
      )
      
      patient_address = (
        sheet.cell(
          row=sheet.max_row - i,
          column=5
        ).value
      )

      patient_age_data = (
        str(
          sheet.cell(
            row=sheet.max_row - i,
            column=3
          ).value
        )
      )

      if patient_age_data == "-" :
        patient_age = ""
      elif "以上" in patient_age_data :
        patient_age = patient_age_data.replace("以上", "歳以上")
      elif (patient_age_data == "園児") or ("未満" in patient_age_data) :
        patient_age = patient_age_data
      else :
        patient_age = patient_age_data + "代"

      patient_gender_data = (
        sheet.cell(
          row=sheet.max_row - i,
          column=4
        ).value
      )

      if patient_gender_data == "-" :
        patient_gender = ""
      else :
        patient_gender = patient_gender_data

      # 退院日
      left_hospital = (
        (
          sheet.cell(
            row=sheet.max_row - i,
            column=6
          ).value
        ) if (
          sheet.cell(
            row=sheet.max_row - i,
            column=6
          ).value
        ) != None else None
      )

      patients_data.append(
        {
          "No": patient_number,
          "リリース日": shorted_date + "T08:00:00.000Z",
          "居住地": patient_address,
          "年代と性別": patient_age + patient_gender,
          "退院": left_hospital,
          "date": shorted_date,
        }
      )

    return patients_data

  def get_patients_summary(self, patients):
    patient_count = 0

    patients_summary = []

    patients_sorted = sorted(
        patients, 
        key=lambda x: datetime.strptime(x["リリース日"], "%Y-%m-%dT%H:%M:%S.000Z")
    )

    for i in range(len(patients_sorted)) :
      date = datetime.strptime(
        patients_sorted[i]["リリース日"], "%Y-%m-%dT%H:%M:%S.000Z"
      )

      if i == 0 :
        recent_patient_date = date

      if (recent_patient_date != date) :
        patients_summary.append(
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

      if i + 1 == len(patients) :
        patients_summary.append(
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
          patients_summary.append(
            {
              "日付": str(zero_date.year).zfill(4) + "-" + str(zero_date.month).zfill(2) + "-" + str(zero_date.day).zfill(2) + "T08:00:00.000Z",
              "小計": 0
            }
          )

      recent_patient_date = date

    return patients_summary

  def get_today_inspctions_summary(self, sheet):
    return {
      "attr": "検査実施人数",
      "value": sheet.cell(row=1, column=2).value,
      "children": [
        {
          "attr": "陽性患者数",
          "value": sheet.cell(row=1, column=3).value,
          "children": [
            {
              "attr": "入院中・入院調整中",
              "value": sheet.cell(row=1, column=5).value
            },
            {
              "attr": "宿泊施設",
              "value": sheet.cell(row=1, column=7).value
            },
            {
              "attr": "自宅療養",
              "value": sheet.cell(row=1, column=8).value
            },
            {
              "attr": "死亡",
              "value": sheet.cell(row=1, column=9).value
            },
            {
              "attr": "退院・解除",
              "value": sheet.cell(row=1, column=4).value
            },
          ]
        }
      ],
      "last_update": None
    }

  def get_inspctions_summary(self, sheet):
    inspections_summary = []
    last_inspected = 0

    for i in range(sheet.max_row):
      inspctions_date = (
        sheet.cell(
          row=sheet.max_row - i,
          column=1
        ).value + timedelta(hours=8)
      )

      inspected = (
        sheet.cell(
          row=sheet.max_row - i,
          column=2
        ).value
      )

      if inspected != None:
        inspections_summary.append(
          {
            "日付": 
              str(inspctions_date.year).zfill(4) + "-" + 
              str(inspctions_date.month).zfill(2) + "-" +
              str(inspctions_date.day).zfill(2) + "T08:00:00.000Z",
            "小計": inspected - last_inspected 
          }
        )

      if inspected != None:
        last_inspected = inspected

    del inspections_summary[0]

    return inspections_summary

class mail_manager:
  def __init__(self, token):
    self.token = token

  def get_message_list(self, addresses: list, date, max_results: int):
    query_addresses = "from:"
    for i in range(len(addresses)):
      if i >= 1:
        query_addresses += "OR from:"
      query_addresses += addresses[i] + " "

    msg_list = gmailrestwrapper.get_message_list(
      self.token,
      query=(
        query_addresses + 
        " after:" +
        str(date.year).zfill(4) + "/" + str(date.month).zfill(2) + "/" + str(date.day).zfill(2) +
        " before:" +
        str(date.year).zfill(4) + "/" + str(date.month).zfill(2) + "/" + str(date.day + 1).zfill(2)
      ),
      maxResults=max_results
    )

    return (
      msg_list["messages"]
      if "messages" in msg_list.keys()
      else None
    )

  def get_message(self, msg_id):
    msg = gmailrestwrapper.get_message(self.token, msg_id)
    date = (
      datetime.utcfromtimestamp(
        int(int(msg["internalDate"]) / 1000)
      ) + timedelta(hours=9)
    )

    return {
      "data": msg,
      "date": str(date)
    }



def __main__():
  data_gen = data_generator()

  param = {
    "client_id": (sys.argv)[1],
    "client_secret": (sys.argv)[2],
    "refresh_token": (sys.argv)[3],
    "grant_type": "refresh_token"
  }

  # トークンを生成
  token = gmailrestwrapper.get_token(param)
  print("a")
  print(token)

  mailman = mail_manager(token)

  dt = datetime.now()

  hoge = True
  while hoge:
    print(dt)
    print("b")
    msg_list = mailman.get_message_list(
      addresses=[
        (sys.argv)[4],
        (sys.argv)[5]
      ],
      date=dt,
      max_results=20
    )

    last_date = None
  
    if msg_list != None:
      for msg_id in msg_list:
        print("c")
        print(msg_id["id"])
        msg = (
            mailman.get_message(
            msg_id["id"]
          )
        )

        for payload in msg["data"]["payload"]["parts"]:
          date = datetime.strptime(msg["date"], "%Y-%m-%d %H:%M:%S")

          if last_date == None:
            last_date = date
          elif date >= last_date:
            last_date = date

          if (
            ("data.xlsx" in payload["filename"])
            and
            (len(payload["filename"]) == 17)
          ):
            filename = payload["filename"]
            print("x")
            spread_data = (
              base64.urlsafe_b64decode(
                (
                  json.loads(
                    gmailrestwrapper.get_attachment(
                      token,
                      msg_id["id"],
                      payload["body"]["attachmentId"]
                    )
                  )
                )["data"].encode("ascii")
              )
            )
            last_update = (
              str(date.year).zfill(4) + "/" +
              str(date.month).zfill(2) + "/" +
              str(date.day).zfill(2) + " " +
              str(date.hour).zfill(2) + ":" +
              str(date.minute).zfill(2)
            )
            hoge = False
            break
        else:
          continue
        break
    else:
      dt = dt + timedelta(days=-1)

  spreadsheet = openpyxl.load_workbook(
    filename=BytesIO(spread_data)
  )

  patients_sheet = spreadsheet['陽性者の属性']
  pcr_sheet = spreadsheet['PCR検査件数']
  news_sheet = spreadsheet['最新の情報']

  patients_data = data_gen.get_patients_data(patients_sheet)

  patients_summary = data_gen.get_patients_summary(patients_data)

  main_summary = data_gen.get_today_inspctions_summary(pcr_sheet)

  inspections_summary = data_gen.get_inspctions_summary(pcr_sheet)

  with open('./data/patients.json', 'w') as f:
    json.dump(
      {
        "data": patients_data,
        "last_update": last_update
      },
      f, indent=4, ensure_ascii=False)

  with open('./data/patients_summary.json', 'w') as f:
    json.dump(
      {
        "data": patients_summary,
        "last_update": last_update
      },
      f, indent=4, ensure_ascii=False)

  with open('./data/main_summary.json', 'w') as f:
    main_summary["last_update"] = last_update
    json.dump(
      main_summary,
      f, indent=4, ensure_ascii=False)

  with open('./data/inspections_summary.json', 'w') as f:
    json.dump(
      {
        "data": inspections_summary,
        "last_update": last_update
      },
      f, indent=4, ensure_ascii=False)

  with open('./data/last_update.json', 'w') as f:
    json.dump(
      {
        "last_update": last_update
      },
      f, indent=4, ensure_ascii=False
    )

  
  

__main__()