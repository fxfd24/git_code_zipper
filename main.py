import requests
import json
import os
import re
import zipfile
import pandas as pd

config = open('config.txt')
for line in config:
    TOKEN = line

try:
    df = pd.read_excel('temp.xlsx')
except:
    data = {'sms_id': [], 'file_name': [], 'count_files': [], 'emoji': []}
    df = pd.DataFrame(data)

def get_updates(offset=0):
    url = f"https://api.telegram.org/bot{TOKEN}/getUpdates"
    params = {"offset": offset, "timeout": 30, "allowed_updates": ['message', 'message_reaction_updated']}
    response = requests.get(url, params=params)
    return response.json()

def is_zip_file(file_name):
    return re.search(r'\.zip$', file_name, re.IGNORECASE) is not None

def download_file(file_id, file_name):
    url = f"https://api.telegram.org/bot{TOKEN}/getFile"
    params = {"file_id": file_id}
    response = requests.get(url, params=params)
    file_info = response.json()
    file_path = file_info["result"]["file_path"]
    url = f"https://api.telegram.org/file/bot{TOKEN}/{file_path}"
    response = requests.get(url, stream=True)
    if is_zip_file(file_name):
        with open(file_name, 'wb') as f:
            for chunk in response.iter_content(1024):
                f.write(chunk)
        # print(f"File downloaded: {file_name}")
        count_files = list_files_in_zip(file_name)
        # print(f'count_files : {count_files}')
        return count_files
    else:
        print(f"Skipping file: {file_name} (not a zip archive)")

    

def list_files_in_zip(zip_file_name):
    with zipfile.ZipFile(zip_file_name, 'r') as zip_file:
        # print(f"Files in archive {zip_file_name}:")
        count = 0
        for file in zip_file.namelist():
            if 'constructionInfo' not in str(file):
                # print(f"- {file}")
                count += 1
        return count

offset = 0
while True:
    updates = get_updates(offset)
    # print(updates)
    if updates["result"]:
        for update in updates["result"]:
            if "message" in update:
                if "document" in update["message"]:
                    sms_id = update["message"]["message_id"]
                    file_id = update["message"]["document"]["file_id"]
                    file_name = update["message"]["document"]["file_name"]           
                    count_files = download_file(file_id, file_name)
                    emoji = '⏳'
                    df.loc[len(df)] = [sms_id, file_name, count_files, emoji]
                    print(f'sms_id : {sms_id}, file_name : {file_name}, count_files : {count_files}')
            
            if "message_reaction" in update:
                sms_id = update["message_reaction"]["message_id"]
                # message_reaction = update["message_reaction"]["new_reaction"]
                if len(update["message_reaction"]["new_reaction"]) != 0:
                    emoji = ''
                    for i in range(len(update["message_reaction"]["new_reaction"])):
                        emoji += (update["message_reaction"]["new_reaction"][i]["emoji"])
                    print(f'sms_id : {sms_id}, message_reaction : {emoji}')
                else:
                    emoji = '⏳'
                    print(f'sms_id : {sms_id}, message_reaction : {emoji}')


                # row_index = df[df['sms_id'] == sms_id]
                # print(df['sms_id'].tolist())
                list_sms_id = df['sms_id'].tolist()
                if sms_id in list_sms_id:
                    row_index = list_sms_id.index(sms_id)
                    # print(row_index)
                    df.loc[row_index, 'emoji'] = emoji
            
            offset = update["update_id"] + 1

            df.to_excel('temp.xlsx', index= False )