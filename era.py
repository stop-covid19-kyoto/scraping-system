import re
import datetime

# date = "令和2年9月15日"

# 引数には文字列を入れるのじゃぞ
def reiwa_to_datetime(date):
  # 元号を判定
  if "令和" in date[0:2] :
    # 元年の場合、year を 1 に設定
    if "元" in date[2:3]:
      year = 1
      month = int(re.search(r'(\d{1,2}[月\./])', date).group().replace("月", ""))
      day = int(re.search(r'(\d{1,2}[日\./])', date).group().replace("日", ""))

    # 年が数字の場合の処理
    elif int(re.search(r'\d+', date).group()):
      year = int(re.search(r'(\d{1,2}[年\./])', date).group().replace("年", ""))
      month = int(re.search(r'(\d{1,2}[月\./])', date).group().replace("月", ""))
      day = int(re.search(r'(\d{1,2}[日\./])', date).group().replace("日", ""))
  
    # 結果を出力
    output_date = datetime.datetime(2019 + year - 1, month, day)
    return output_date
