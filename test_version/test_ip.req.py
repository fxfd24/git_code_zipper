from telethon import TelegramClient, events
import os
import datetime
from telethon.tl.types import UpdateNewChannelMessage, UpdateEditChannelMessage, MessageMediaDocument, UpdateBotMessageReaction
import shutil
import re
import zipfile
import rarfile
import patoolib
import pandas as pd


import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from gspread_formatting.dataframe import format_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import textdistance

from collections import OrderedDict

api_id = '27857864'
api_hash = '3224d8685a163fbaafc782f30e95937f'
bot_token = '7372987398:AAFpA4weK_NXtCjrn5gnmeFp3lVwE_RVw0s'


googleJsonPath = "google_secret/client_secret.json"
creds_sheets = ServiceAccountCredentials.from_json_keyfile_name(googleJsonPath, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"])
gs = gspread.authorize(creds_sheets)
sh = gs.open('Botick_memory')
worksheet = sh.sheet1
# worksheet.clear()

data = {'sms_id': [], 'Имя файла': [], 'Номер наряда': [], 'Пациент': [], 'Врач': [], 'Тип': [], 'Перевыпуск': [], 'Кол-во файлов': [], 'emoji': [], 'Техник': [], 'Характеристика': [],'Дата появления в чате': [],'Готова фактически': [],'Комментарий':[],'Number - ImplantLibraryEntryDescriptor':[]}
df_ = pd.DataFrame(data)

columns = ['sms_id', 'Имя файла', 'Номер наряда', 'Пациент', 'Врач', 'Тип', 'Перевыпуск', 'Кол-во файлов', 'emoji', 'Техник', 'Характеристика', 
							'Дата появления в чате', 'Готова фактически', 'Комментарий', 
							'Number - ImplantLibraryEntryDescriptor']
df = pd.DataFrame(worksheet.get_all_records(expected_headers = columns)) #получить
#df_ = 
set_with_dataframe(worksheet, df_)
format_with_dataframe(worksheet, df_, include_column_header=True) #записать
print('Google Excel Connected Successfully')
# print(worksheet.get_all_records())

emoji_list = {'❤️':'Работа готова','🤬':'Брак конструкции (ошибка в модели)','💔':'Брак модели (проблема с печатью)','😡':'Частично в работе','🤔':'Решаем проблему',}
client = TelegramClient('bot', api_id, api_hash).start(bot_token=bot_token)
print('TelegramClient Connected Successfully')
# Словарь для хранения информации о реакциях на сообщения
message_reactions = {}

def is_zip_file(file_name):
    return re.search(r'\.zip$', file_name, re.IGNORECASE) is not None


def list_files_in_zip(file_name):
    with zipfile.ZipFile(file_name, 'r') as zip_file:
        print(f"Files in archive {zip_file.namelist()}:")
        count = 0
        for file in zip_file.namelist():
            if 'constructionInfo' not in str(file) and '.stl' in str(file):
                # print(f"- {file}")
                count += 1
        return count

def extract_constructionInfo_from_zip(file_name):
    ci = False
    with zipfile.ZipFile(file_name, 'r') as zip_file:
        for file in zip_file.namelist():
            if 'constructionInfo' in str(file):
                print(f'Нашел constructionInfo в {file_name}')
                ci = True
                try:
                    with zip_file.open(file) as f:
                        content = f.read().decode('utf-8')
                        # print(content)
                        numbers = re.findall(r'<Number>(.*?)</Number>', content)
                        descriptors = re.findall(r'<ImplantLibraryEntryDescriptor>(.*?)</ImplantLibraryEntryDescriptor>', content)
                        str_n_d = ''
                        if not descriptors:
                            for i in range(len(numbers)):
                                # print(descriptors)
                                descriptors.append('Не найден')
                        if numbers and descriptors:
                            for i, (number, descriptor) in enumerate(zip(numbers, descriptors), start=1):
                                str_n_d += f'{i} пара: Number={number}, Descriptor={descriptor}\n'
                                print(f"{i} пара: Number={number}, Descriptor={descriptor}")
                        else:
                            str_n_d += 'NoneInfo'
                        return str_n_d
                except:
                    str_n_d = 'ConstructionInfoOpeningErrorNone'
                    return str_n_d
    if ci == False:
        str_n_d = 'NoneConstructionInfo'
        return str_n_d

def is_rar_file(file_name):
    return re.search(r'\.rar$', file_name, re.IGNORECASE) is not None

def list_files_in_rar(file_name):
    with rarfile.RarFile(file_name, 'r') as rar_ref:
        print(f"Files in archive {rar_ref.namelist()}:")
        count = 0
        for file in rar_ref.namelist():
            if 'constructionInfo' not in str(file) and '.stl' in str(file):
                # print(f"- {file}")
                count += 1
        return count

def extract_constructionInfo_from_rar(file_name):
    ci = False
    with rarfile.RarFile(file_name, 'r') as rar_ref:
        for file in rar_ref.namelist():
            if 'constructionInfo' in str(file):
                print(f'Нашел constructionInfo в {file_name}')
                ci = True
                shutil.rmtree('temp')
                patoolib.extract_archive(file_name, outdir="temp")
                try:
                    with open(f"temp/{file}", "r", encoding="utf-8") as file:
                        content = file.read()
                        numbers = re.findall(r'<Number>(.*?)</Number>', content)
                        descriptors = re.findall(r'<ImplantLibraryEntryDescriptor>(.*?)</ImplantLibraryEntryDescriptor>', content)
                        str_n_d = ''
                        if not descriptors:
                            for i in range(len(numbers)):
                                # print(descriptors)
                                descriptors.append('Не найден')
                        if numbers and descriptors:
                            for i, (number, descriptor) in enumerate(zip(numbers, descriptors), start=1):
                                str_n_d += f'{i} пара: Number={number}, Descriptor={descriptor}\n'
                                print(f"{i} пара: Number={number}, Descriptor={descriptor}")
                        else:
                            str_n_d += 'NoneInfo'
                        return str_n_d
                except:
                    str_n_d = 'ConstructionInfoOpeningErrorNone'
                    return str_n_d

    if ci == False:
        str_n_d = 'NoneConstructionInfo'
        return str_n_d

def is_stl_file(file_name):
    return re.search(r'\.stl$', file_name, re.IGNORECASE) is not None


@client.on(events.Raw())
async def handler(event):
	global message_reactions
	# print(f"\nEvent: {event}\n")
	if isinstance(event, UpdateNewChannelMessage):
		print(f"New channel message received: {event.message}")
		if event.message.media and hasattr(event.message.media, 'document'):
			document = event.message.media.document
			message_id = event.message.id
			caption = event.message.message
			document = event.message.media.document
			file_name = document.attributes[0].file_name
			file_id = document.id

			print(f"Document received: Title: {document.attributes[0].file_name}, ID: {document.id}, Message: {event.message.message}")
			print(f"sms_id : {message_id}, file_name: {file_name}, Комментарий: {caption}")
			
			

			# Путь, куда будет сохранен документ
			path = os.path.join('downloads', file_name)

			if is_zip_file(path):
				# Скачиваем документ
				await client.download_media(document, file=path)
				print(f"Document downloaded and saved to {path}")
				#Работаем со скачанным файлом zip
				count_files = list_files_in_zip(path)
				pares = extract_constructionInfo_from_zip(path)
				print(f'message_id: {message_id}, count_files : {count_files}, pares : {pares}')
				os.remove(path)
				print(f'Удалил {path}')
			elif is_rar_file(path):
				# Скачиваем документ
				await client.download_media(document, file=path)
				print(f"Document downloaded and saved to {path}")
				#Работаем со скачанным файлом zip
				count_files = list_files_in_rar(path)
				pares = extract_constructionInfo_from_rar(path)
				print(f'message_id: {message_id}, count_files : {count_files}, pares : {pares}')
				os.remove(path)
			elif is_stl_file(path):
				# Скачиваем документ
				await client.download_media(document, file=path)
				print(f"Document downloaded and saved to {path}")
				#Работаем со скачанным файлом stl
				count_files = 1
				pares = '.stl, NoneConstructionInfo'
				print(f'count_files : {count_files}, pares : {pares}')
				os.remove(path)
			else:
				print(f"Skipping file: {file_name} (not a zip or archive)")

			per = 'none'
			if 'перевыпуск' in file_name.lower():
				per = 'Перевыпуск'
			if file_name[0].isdigit():
				number = file_name[0:8]
				file_name_unix = file_name[8::]
				result = [x for x in re.split(r'[_-]', file_name_unix) if x]
				if len(result) == 3:
					pac = result[0]
					med = result[1]
					tip = result[2][:-4]
				elif len(result) > 3:
					pac = result[0]
					med = result[1]
					tip = result[2]
				else:
					pac, med, tip = 'none','none','none'
			else:
				number, pac, med, tip = 'none','none','none','none'
			if caption:
				pass
			else:
				caption = 'Отсутсвует'     
			emoji = '⏳'
			date_unix = datetime.datetime.now()
			# Прибавление 3 часов
			date_unix_plus_3_hours = date_unix + datetime.timedelta(hours=3)
			# Преобразование в строковый формат
			date_normal = date_unix_plus_3_hours.strftime("%d-%m-%Y %H:%M")
			#_____первая____запись_____документа_____
			new_row = [message_id, file_name, number, pac, med, tip, per, count_files, emoji, 'none', 'none', date_normal,'none',caption, pares]
			columns = ['sms_id', 'Имя файла', 'Номер наряда', 'Пациент', 'Врач', 'Тип', 'Перевыпуск', 'Кол-во файлов', 'emoji', 'Техник', 'Характеристика', 
							'Дата появления в чате', 'Готова фактически', 'Комментарий', 
							'Number - ImplantLibraryEntryDescriptor']
			df = pd.DataFrame(worksheet.get_all_records(expected_headers = columns)) #получить
			df_ = pd.DataFrame([new_row], columns=columns)
			updated_df = pd.concat([df, df_], ignore_index=True)
			set_with_dataframe(worksheet, updated_df)
			format_with_dataframe(worksheet, updated_df, include_column_header=True) #записать
			print('Записал')

	elif isinstance(event, UpdateEditChannelMessage):
		message_id = event.message.id
		# print(f"Channel message edited: {event}")
		# Запоминаем id сообщения для дальнейшего использования при обработке реакций
		message_reactions[message_id] = {'old_reactions': [], 'new_reactions': []}
	elif isinstance(event, UpdateBotMessageReaction):
		print(event)
		char = ''
		date_finish = ''
		date_unix = datetime.datetime.now()
		# Прибавление 3 часов
		date_unix_plus_3_hours = date_unix + datetime.timedelta(hours=3)
		# Преобразование в строковый формат
		date_normal = date_unix_plus_3_hours.strftime("%d-%m-%Y %H:%M")
		sms_id = event.msg_id
		user_tg = event.actor.user_id
		# print(user_tg)
		emoji = ''
		list_old_e = []
		message_id = event.msg_id

		if message_id in message_reactions:
			# print(f'{[reaction.emoticon for reaction in event.old_reactions]}')
			# print(f'{[reaction.emoticon for reaction in event.new_reactions]}')
			old_reactions = [reaction.emoticon for reaction in event.old_reactions]
			new_reactions = [reaction.emoticon for reaction in event.new_reactions]
			# Извлекаем символы реакций и добавляем их в список
			# old_reactions.extend([reaction.emoticon for reaction in event.old_reactions])
			# new_reactions.extend([reaction.emoticon for reaction in event.new_reactions])
			message_reactions[message_id]['old_reactions'] = old_reactions
			message_reactions[message_id]['new_reactions'] = new_reactions
			print(f"Reactions for message ID {message_id}:")
			print(f"Old reactions: {old_reactions}")
			print(f"New reactions: {new_reactions}")
			# Вывод текущей реакции на сообщении
			# current_reactions = list(set(new_reactions) - set(old_reactions))
			# print(f"Current reaction on message ID {message_id}: {current_reactions}")
			# if len(current_reactions) != 0:
			if len(new_reactions) != 0:
				for i in range(len(old_reactions)):
					list_old_e.append(old_reactions[i])
				for i in range(len(new_reactions)):
					if new_reactions[i] not in list_old_e:
						# print(update["message_reaction"]["new_reaction"][i]["emoji"])
						if new_reactions[i] == '❤':
							char += f'❤️ - {emoji_list['❤️']}'
							date_finish += f'{date_normal}'
							# print(emoji_list['❤️'])
						elif new_reactions[i] == '🤬':
							char += f'🤬 - {emoji_list['🤬']}'
							# print(emoji_list['🤬'])
						elif new_reactions[i] == '💔':
							char += f'💔 - {emoji_list['💔']}'
							# print(emoji_list['💔'])
						elif new_reactions[i] == '😡':
							char += f'😡 - {emoji_list['😡']}'
							# print(emoji_list['😡'])
						elif new_reactions[i] == '🤔':
							char += f'🤔 - {emoji_list['🤔']}'
							# print(emoji_list['🤔'])

				for i in range(len(new_reactions)):
					emoji += new_reactions[i]
					print(f'sms_id : {sms_id}, message_reaction : {emoji}, char :{char}')
			else:
				emoji = '⏳'
				print(f'sms_id : {sms_id}, message_reaction : {emoji}, char :{char}')

			# Читаем таблицу
			columns = ['sms_id', 'Имя файла', 'Номер наряда', 'Пациент', 'Врач', 'Тип', 'Перевыпуск', 'Кол-во файлов', 'emoji', 'Техник', 'Характеристика', 
							'Дата появления в чате', 'Готова фактически', 'Комментарий', 
							'Number - ImplantLibraryEntryDescriptor']
			df = pd.DataFrame(worksheet.get_all_records(expected_headers = columns))
			# df = pd.read_excel('temp.xlsx')
			list_sms_id = df['sms_id'].tolist()
			if sms_id in list_sms_id:
				row_index = list_sms_id.index(sms_id)
				df.loc[row_index, 'emoji'] = emoji
				if char:
					# print(f'{char} - {update["message_reaction"]["user"]["username"]} - {date_normal}')
					if df['Характеристика'][row_index] == 'none':
						df.loc[row_index, 'Характеристика'] = f'{char} from user_id: {user_tg} - {date_normal}'
					else:
						df.loc[row_index, 'Характеристика'] = f'{df['Характеристика'][row_index]}\n{char} from user_id: {user_tg} - {date_normal}'
				if df['Техник'][row_index] == 'none':
					df.loc[row_index, 'Техник'] = f'user_id: {user_tg}'
				if date_finish:
					df.loc[row_index, 'Готова фактически'] = f'{date_finish}'

			set_with_dataframe(worksheet, df)
			format_with_dataframe(worksheet, df, include_column_header=True) #записать
			print('Записал Реакцию')
			# df.to_excel('temp.xlsx', index= False )

client.start()
client.run_until_disconnected()
