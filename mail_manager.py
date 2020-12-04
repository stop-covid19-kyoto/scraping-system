from datetime import datetime, timedelta, timezone
import gmailrestwrapper

class mail_manager:
  def __init__(self, token):
    self.token = token

  def get_message_list(self, addresses: list, date, max_results: int):
    query_addresses = "from:"

    date = date
    before_date = date + timedelta(days=1)

    date_unixtime = date.strftime('%s')
    before_date_unixtime = before_date.strftime('%s')

    for i in range(len(addresses)):
      if i >= 1:
        query_addresses += "OR from:"
      query_addresses += addresses[i] + " "

    query=(
        "NOT in:draft " +
        "has:attachment " +
        query_addresses + 
        " after:" +
        date_unixtime +
        " before:" +
        before_date_unixtime
      )

    msg_list = gmailrestwrapper.get_message_list(
      self.token,
      query=query,
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
      )
    )
    
    return {
      "data": msg,
      "date": str(date + timedelta(hours=9))
    }
