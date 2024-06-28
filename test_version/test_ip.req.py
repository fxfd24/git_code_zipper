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
import random


import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from gspread_formatting.dataframe import format_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import textdistance

from collections import OrderedDict

from telethon.tl.functions.messages import SendMediaRequest
from telethon.tl.types import InputMediaUploadedPhoto, MessageEntityBold, MessageEntityUnderline
from telethon.tl.functions.channels import GetParticipantsRequest
from telethon.tl.types import ChannelParticipantsSearch

api_id = '27857864'
api_hash = '3224d8685a163fbaafc782f30e95937f'
bot_token = '7372987398:AAFpA4weK_NXtCjrn5gnmeFp3lVwE_RVw0s'


googleJsonPath = "google_secret/client_secret.json"
creds_sheets = ServiceAccountCredentials.from_json_keyfile_name(googleJsonPath, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"])
gs = gspread.authorize(creds_sheets)
# sh = gs.open('Botick_memory')
sh = gs.open('test_bot_click')
worksheet = sh.sheet1
# worksheet.clear()

data = {'Характеристика': [],'Номер наряда': [], 'Пациент': [], 'Врач': [],  'Тип': [], 'Готова фактически': [], 'Дата появления в чате': [],'Количество': [], 'Цвет': [], 'Комментарий':[], 'Кол-во файлов': [],'Перевыпуск': [], 'Исполнитель': [], 'Техник': [],  'emoji': [], 'Имя файла': [], 'Number - ImplantLibraryEntryDescriptor':[],  'sms_id': []}
df_ = pd.DataFrame(data)

columns = ['Характеристика','Номер наряда','Пациент','Врач','Тип','Готова фактически',  'Дата появления в чате','Количество','Цвет','Комментарий','Кол-во файлов','Перевыпуск', 'Исполнитель', 'Техник', 'emoji', 'Имя файла','Number - ImplantLibraryEntryDescriptor','sms_id']

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


# Функция для отправки сообщения с картинкой
async def send_image_with_text(client, chat_id, image_path, bold_text, underlined_text, reply_to):
	message_text = f"<b>{bold_text}</b>\n<u>{underlined_text}</u>"
	input_file = await client.upload_file(image_path)
	await client.send_file(chat_id, file=input_file, caption=message_text, parse_mode='HTML', reply_to=reply_to)

# Функция для получения участников группы
async def get_group_participants(client, group_id):
    participants = await client(GetParticipantsRequest(
        channel=group_id,
        filter=ChannelParticipantsSearch(''),
        offset=0,
        limit=100,
        hash=0
    ))
    return {user.id: user.username for user in participants.users}



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
			user_id_ = event.message.from_id.user_id
			all_users = await get_group_participants(client,-1002154104395)
			user_tg = all_users[user_id_]
			if user_tg is None:
				user_tg = user_id_
			print(f'user_tg>>>{user_tg}')

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

			
			if file_name[0].isdigit():
				number = file_name[0:8]
				file_name_unix = file_name[8::]
				result = [x for x in re.split(r'[_-]', file_name_unix) if x]
				print(result)
				if len(result) == 5:
					med = result[0]
					pac = result[1]
					tip = result[2]
					col = result[3]
					color = result[4][:-4]
				elif len(result) == 4:
					med = result[0]
					pac = result[1]
					tip = result[2]
					col = result[3][:-4]
					color = 'нет'
				else:
					# тут бот должен отправить картинку с текстом, в котором одна строка жирная, а другая подчерктуна
					pac, med, tip ,col ,color= 'none','none','none','none','none'
					# Отправка картинки с текстом
					random_number = random.randint(1, 11)
					print(random_number)
					image_path = f'img/mem{random_number}.jpg'
					bold_text = "Некорректное заполнение, переделываем архив \n"
					underlined_text = "НОМЕР - ПАЦИЕНТ - ВРАЧ - ТИП - КОЛИЧЕСТВО- ЦВЕТ"
					await send_image_with_text(client, event.message.chat_id, image_path, bold_text, underlined_text,event.message.id)

			else:
				number, pac, med, tip ,col ,color= 'none','none','none','none','none','none'
				# Отправка картинки с текстом
				random_number = random.randint(1, 11)
				print(random_number)
				image_path = f'img/mem{random_number}.jpg'
				bold_text = "Некорректное заполнение, переделываем архив \n"
				underlined_text = "НОМЕР - ПАЦИЕНТ - ВРАЧ - ТИП - КОЛИЧЕСТВО- ЦВЕТ"
				await send_image_with_text(client, event.message.chat_id, image_path, bold_text, underlined_text,event.message.id)
			if caption:
				pass
			else:
				caption = 'Отсутсвует'
			per = 'Нет'
			if 'перевыпуск' in caption.lower():
				per = 'Перевыпуск'     
			emoji = '⏳'
			date_unix = datetime.datetime.now()
			# Прибавление 3 часов
			date_unix_plus_3_hours = date_unix + datetime.timedelta(hours=3)
			# Преобразование в строковый формат
			date_normal = date_unix_plus_3_hours.strftime("%d-%m-%Y %H:%M")
			#_____первая____запись_____документа_____
			# new_row = [message_id, file_name, number, pac, med, tip,  emoji, 'none', 'none', date_normal,'none', pares]

			new_row = ['В работе',number,med, pac,  tip, 'none', date_normal, col,color, caption, count_files, per, 'none',user_tg, emoji,  file_name, pares,message_id]
			# columns = ['emoji','Номер наряда','Врач','Пациент','Тип','Количество','Цвет','Комментарий','Кол-во файлов','Перевыпуск','Техник','Дата появления в чате','Готова фактически','Характеристика','Имя файла','Number - ImplantLibraryEntryDescriptor','sms_id']
			columns = ['Характеристика','Номер наряда','Пациент','Врач','Тип','Готова фактически',  'Дата появления в чате','Количество','Цвет','Комментарий','Кол-во файлов','Перевыпуск','Исполнитель','Техник', 'emoji', 'Имя файла','Number - ImplantLibraryEntryDescriptor','sms_id']

			df = pd.DataFrame(worksheet.get_all_records(expected_headers = columns)) #получить
			df_ = pd.DataFrame([new_row], columns=columns)
			updated_df = pd.concat([df, df_], ignore_index=True)
			set_with_dataframe(worksheet, updated_df)
			format_with_dataframe(worksheet, updated_df, include_column_header=True) #записать
			print('Записал')

	elif isinstance(event, UpdateEditChannelMessage):
		# print(event)
		message_id = event.message.id
		# print(f"Channel message edited: {event}")
		# Запоминаем id сообщения для дальнейшего использования при обработке реакций
		message_reactions[message_id] = {'old_reactions': [], 'new_reactions': []}
	elif isinstance(event, UpdateBotMessageReaction):
		print(event)
		all_users = await get_group_participants(client,-1002154104395)
		print(all_users)
		char = ''
		date_finish = ''
		date_unix = datetime.datetime.now()
		# Прибавление 3 часов
		date_unix_plus_3_hours = date_unix + datetime.timedelta(hours=3)
		# Преобразование в строковый формат
		date_normal = date_unix_plus_3_hours.strftime("%d-%m-%Y %H:%M")
		sms_id = event.msg_id
		user_id_ = event.actor.user_id
		# user_id_ = 1223812779 #челик без никнейма
		user_tg = all_users[user_id_]
		if user_tg is None:
			user_tg = user_id_
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
			columns = ['Характеристика','Номер наряда','Пациент','Врач','Тип','Готова фактически',  'Дата появления в чате','Количество','Цвет','Комментарий','Кол-во файлов','Перевыпуск','Исполнитель','Техник', 'emoji', 'Имя файла','Number - ImplantLibraryEntryDescriptor','sms_id']

			df = pd.DataFrame(worksheet.get_all_records(expected_headers = columns))
			list_sms_id = df['sms_id'].tolist()
			if sms_id in list_sms_id:
				row_index = list_sms_id.index(sms_id)
				df.loc[row_index, 'emoji'] = emoji
				if char:
					# print(f'{char} - {update["message_reaction"]["user"]["username"]} - {date_normal}')
					if df['Характеристика'][row_index] == 'В работе':
						df.loc[row_index, 'Характеристика'] = f'{char} from: {user_tg} - {date_normal}'
					else:
						df.loc[row_index, 'Характеристика'] = f'{df['Характеристика'][row_index]}\n{char} from: {user_tg} - {date_normal}'
				if df['Исполнитель'][row_index] == 'none':
					df.loc[row_index, 'Исполнитель'] = f'{user_tg}'
				if date_finish:
					df.loc[row_index, 'Готова фактически'] = f'{date_finish}'

			set_with_dataframe(worksheet, df)
			format_with_dataframe(worksheet, df, include_column_header=True) #записать
			print('Записал Реакцию')


client.start()
client.run_until_disconnected()
