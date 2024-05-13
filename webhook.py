import requests
import json

config = open('config.txt')
for line in config:
    TOKEN = line

def test_webhook():
    url = f"https://api.telegram.org/bot{TOKEN}/getWebhookInfo"
    response = requests.get(url)
    return response.json()

print(test_webhook())

def set_webhook():
    url = f"https://api.telegram.org/bot{TOKEN}/setWebhook"
    data = {'allowed_updates': ['message','message_reaction'] }
    response = requests.post(url, json=data)
    return json.loads(response.text)

# res = set_webhook()
# print(res)

# def delete_webhook():
#     url = f"https://api.telegram.org/bot{TOKEN}/deleteWebhook"
#     response = requests.get(url)
#     return response.json()

# print(delete_webhook())