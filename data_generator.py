from datetime import datetime, timedelta
import era

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

      print(f'{i + 1}件処理しました')

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

  def get_today_inspections_summary(self, sheet):
    inspected = None
    patients = None
    icu = None
    other = None
    in_hospital = None
    in_hotel = None
    in_home = None
    dead = None
    left = None
    adjustment = None

    for i in range(sheet.max_row):
      if (
        (inspected != None) and
        (patients != None) and
        (icu != None) and
        (other != None) and
        (in_hospital != None) and
        (in_hotel != None) and
        (in_home != None) and
        (dead != None) and
        (left != None) and
        (adjustment != None)
      ):
        break

      # 検査実施人数
      inspected = (
        sheet.cell(row=i + 1, column=2).value
        if inspected == None 
        else inspected
      )
      # 陽性者
      patients = (
        sheet.cell(row=i + 1, column=3).value
        if patients == None
        else patients
      )
      # 高度重症病床
      icu = (
        sheet.cell(row=i + 1, column=6).value
        if icu == None
        else icu
      )
      # その他
      other = (
        sheet.cell(row=i + 1, column=7).value
        if other == None
        else other
      )
      # 入院中
      in_hospital = (
        sheet.cell(row=i + 1, column=5).value
        if in_hospital == None
        else in_hospital
      )
      # 施設療養
      in_hotel = (
        sheet.cell(row=i + 1, column=8).value
        if in_hotel == None
        else in_hotel
      )
      # 自宅療養
      in_home = (
        sheet.cell(row=i + 1, column=9).value
        if in_home == None
        else in_home
      )
      # 死亡
      dead = (
        sheet.cell(row=i + 1, column=10).value
        if dead == None
        else dead
      )
      # 退院
      left = (
        sheet.cell(row=i + 1, column=4).value
        if left == None
        else left
      )
      # 調整中
      adjustment = (
        sheet.cell(row=i + 1, column=11).value
        if adjustment == None
        else adjustment
      )

    return {
      "attr": "検査実施人数",
      "value": inspected if inspected != None else 0,
      "children": [
        {
          "attr": "陽性患者数",
          "value": patients if patients != None else 0,
          "children": [
            {
              "attr": "入院中・入院調整中",
              "value": in_hospital if in_hospital != None else 0
            },
            {
              "attr": "高度重症病床",
              "value": icu if icu != None else 0
            },
            {
              "attr": "その他",
              "value": other if other != None else 0
            },
            {
              "attr": "宿泊施設",
              "value": in_hotel if in_hotel != None else 0
            },
            {
              "attr": "自宅療養",
              "value": in_home if in_home != None else 0
            },
            {
              "attr": "死亡",
              "value": dead if dead != None else 0
            },
            {
              "attr": "退院・解除",
              "value": left if left != None else 0
            },
            {
              "attr": "調整中",
              "value": adjustment if adjustment != None else 0
            },
          ]
        }
      ],
      "last_update": None
    }

  def get_inspections_summary(self, sheet):
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

    return inspections_summary

  def get_news(self, sheet):
    news_items = []
    for i in range(sheet.max_row):
      date = (
        (
          sheet.cell(
            row=i + 1,
            column=1
          ).value
        ) + timedelta(hours=8)
      )

      news_items.append(
        {
          "date": str(date.year).zfill(4) + "/" + 
                  str(date.month).zfill(2) + "/" +
                  str(date.day).zfill(2),
          "text": sheet.cell(
            row=i + 1,
            column=2
          ).value,
          "url": str(
            sheet.cell(
              row=i + 1,
              column=3
            ).value
          )
        }
      )
    return news_items
