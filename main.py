from datetime import datetime, timedelta
from io import BytesIO
from mail_manager import mail_manager
from data_generator import data_generator

import json
import openpyxl
import gmailrestwrapper
import base64
import sys

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

  print("メールサーバにアクセスします")
  mail_man = mail_manager(token)

  dt = datetime.now()
  dt = dt.replace(hour=0, minute=0, second=0, microsecond=0)
  dt = dt - timedelta(hours=9)

  msg_list = mail_man.get_message_list(
    addresses=[
      (sys.argv)[4],
      (sys.argv)[5]
    ],
    date=dt,
    max_results=20
  )

  last_date = None

  spread_data = None

  print("メッセージを取得しています")
  if msg_list != None:
    for msg_id in msg_list:
      msg = (
        mail_man.get_message(
          msg_id["id"]
        )
      )
      if ("parts" in msg["data"]["payload"].keys()) == True:
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
            break
      else:
        continue
      break
  else:
    dt = dt + timedelta(days=-1)

  if spread_data != None:
    spreadsheet = openpyxl.load_workbook(
      filename=BytesIO(spread_data)
    )

    patients_sheet = spreadsheet['陽性者の属性']
    pcr_sheet = spreadsheet['PCR検査件数']
    news_sheet = spreadsheet['最新の情報']

    print("ニュースを変換中です")
    news = data_gen.get_news(news_sheet)

    print("感染者のデータを変換中です")
    patients_data = data_gen.get_patients_data(patients_sheet)

    print("感染者数のサマリーを変換中です")
    patients_summary = data_gen.get_patients_summary(patients_data)

    print("検査陽性者の属性データを変換中です")
    main_summary = data_gen.get_today_inspections_summary(pcr_sheet)

    print("検査実施数のデータを変換中です")
    inspections_summary = data_gen.get_inspections_summary(pcr_sheet)

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

    with open('./data/news.json', 'w') as f:
      json.dump(
        {
          "newsItems": news
        },
        f, indent=4, ensure_ascii=False
      )

__main__()
