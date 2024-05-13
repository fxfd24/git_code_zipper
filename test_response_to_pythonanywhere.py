import requests

url = 'http://KeyKoal2024.pythonanywhere.com/process_text'  # replace with the actual endpoint URL
text = 'Какой бы первый президент РОссии?'  # replace with the actual text you want to send

response = requests.post(url, data={'text': text})

if response.status_code == 200:
    print(response.text)
else:
    print(f'Request failed with status code {response.status_code}')