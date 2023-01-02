import requests

class Bot (object):
    def send_telegram(self, text: str):
        token = "5939978159:AAF03tgum5GWNgfZ-wBcv0lC6TGT-D3ggeA"
        url = "https://api.telegram.org/bot"
        channel_id = "@RussianMogul"
        url += token
        method = url + "/sendMessage"

        r = requests.post(method, data={
             "chat_id": channel_id,
             "text": text
              })

        if r.status_code != 200:
            raise Exception("post_text error")
