from datetime import datetime, timedelta
import gmailrestwrapper

class mail_manager:
  def __init__(self, token):
    self.token = token

  def get_message_list(self, addresses: list, date, max_results: int):
    query_addresses = "from:"
    date = date - timedelta(days=1)
    before_date = date + timedelta(days=1)
    for i in range(len(addresses)):
      if i >= 1:
        query_addresses += "OR from:"
      query_addresses += addresses[i] + " "

    msg_list = gmailrestwrapper.get_message_list(
      self.token,
      query=(
        "NOT in:draft " +
        "has:attachment " +
        query_addresses + 
        " after:" +
        str(date.year).zfill(4) + "/" + str(date.month).zfill(2) + "/" + str(date.day).zfill(2) +
        " before:" +
        str(before_date.year).zfill(4) + "/" + str(before_date.month).zfill(2) + "/" + str(before_date.day).zfill(2)
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
