import requests
import json

def get_token(client):
  """
  トークンを取得する。

  Parameters
  ----------
  client: dict
    クライアント情報を格納した辞書型オブジェクト

  Returns
  -------
  token: str
    アクセストークン
  """
  req = requests.post(
    "https://accounts.google.com/o/oauth2/token",
    params=client
  )
  token = str(json.loads(req.text)["access_token"])
  return token

def get_message_list(token, query="", maxResults=10):
  """
  トークンを取得する。

  Parameters
  ----------
  client: dict
    クライアント情報を格納した辞書型オブジェクト

  Returns
  -------
  message_list: str
    メッセージのリスト
  """
  head = {
    "Authorization": "Bearer " + token
  }

  param= {
    "maxResults": maxResults,
    "q": query
  }

  req = requests.get(
    "https://www.googleapis.com/gmail/v1/users/me/messages/",
    headers=head,
    params=param
  )

  message_list = dict(json.loads(req.text))
  return message_list

def get_message(token, msgid):
  head = {
    "Authorization": "Bearer " + token
  }

  param= {
    "maxResults": msgid,
  }

  req = requests.get(
    "https://www.googleapis.com/gmail/v1/users/me/messages/" + msgid,
    headers=head,
    params=param
  )

  message = dict(json.loads(req.text))
  return message

def get_attachment(token, msgid="", attachment_id=""):
  head = {
    "Authorization": "Bearer " + token
  }

  param= {
    "messageId": msgid,
    "attachmentID": attachment_id
  }

  req = requests.get(
    "https://www.googleapis.com/gmail/v1/users/me/messages/" + msgid + "/attachments/" + attachment_id,
    headers=head
  )

  attachment = req.text
  return attachment
