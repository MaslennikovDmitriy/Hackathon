#!/usr/bin/env python
import telebot
from telebot import types
import requests
import wget
import numpy as np
import openpyxl as op

import pandas as pd
import sqlite3 as sl
import os

import unicodedata # для игнора регистра


TOKEN = "6703792810:AAHn1tEMb6oeEpSVEocOr4J9j0WIdTGpwqM"
bot = telebot.TeleBot(TOKEN)

all_types = ["text", "audio", "document", "photo", "sticker", "video", "video_note", "voice", "location", "contact", "poll"]



@bot.message_handler(commands=['start', 'hello']) # приветствие 
def send_welcome(message):
    """Отправляет приветствие пользователю после ввода команд /start или /hello в боте."""
    
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    bot.reply_to(message, "Здравствуй, " + message.chat.first_name + "!", parse_mode='html', reply_markup=markup)

@bot.message_handler(commands=['help']) # список команд
def help(message):
    """Показывает все доступные команды в боте после введения команды /help."""
    
    bot.reply_to(message, "/upload_excel - обновить базу данных\n/change_password - смена пароля администратора\n/dataset_info - работа с базой данных")

@bot.message_handler(commands=['upload_excel']) # загрузка таблицы
def upload_excel(message):
    """После ввода команды /upload_excel появляется возможность загрузки excel-таблицы. Предлагается ввести пароль для дальнейшего доступа к загрузке excel-таблицы."""
    
    markup = types.ReplyKeyboardRemove()
    sent = bot.send_message(message.chat.id, 'Введите пароль', reply_markup=markup)
    bot.register_next_step_handler(sent, tab_login)

@bot.message_handler(commands=['change_password']) # смена пароля
def change_password(message):
    """После ввода команды /change_password в боте администратор может поменять пароль. Требует ввод текущего текста пароля для доступа к дальнейшему изменению.
        Важно, что для корректной работы ввода пароля файл password.txt с текстом пароля должен храниться в корневой директории вместе с кодом."""
    
    markup = types.ReplyKeyboardRemove()
    sent = bot.send_message(message.chat.id, 'Введите старый пароль', reply_markup=markup)
    bot.register_next_step_handler(sent, pass_login)


def pass_login(message): # проверка пароля
    """Описывает сценарий ответов бота на введеный пароль (password) после команды /change_password в боте, реагирует на верность/неверность введеного администратором пароля. После правильного предлагает ввести новый текст пароля.
        В случае неверно введеного пароля, бот сообщает об этом и предлагает клавиатуру с кнопками "Выйти" и "Попробовать снова".
        Важно, что для корректной работы ввода пароля файл password.txt с текстом пароля должен храниться в корневой директории вместе с кодом."""
    
    password = pass_reader() # чтение пароля из файла
    if message.text == password:
        markup = types.ReplyKeyboardRemove()
        sent = bot.send_message(message.chat.id, 'Успешный вход. Введите новый пароль')
        bot.register_next_step_handler(sent, pass_writer)
    else:
       markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
       but1 = types.KeyboardButton("Выйти")
       but2 = types.KeyboardButton("Попробовать снова")
       markup.add(but1, but2)
       sent = bot.send_message(message.chat.id, 'Пароль неверен', reply_markup=markup)
       bot.register_next_step_handler(sent, try_again_pass_change)

def tab_login(message): # проверка пароля при попытке загрузки таблицы
    """Предлагает администратору ввести пароль (password) в ответ на /upload_excel. В случае верно введенного пароля бот сообщает об успешном входе и предлагает отправить таблицу в формате excel следующим сообщением.
        В случае, когда неверно введен пароль, бот сообщает об этом и предлагает клавиатуру с кнопками "Выйти" и "Попробовать снова".
        Важно, что для корректной работы ввода пароля файл password.txt с текстом пароля должен храниться в корневой директории вместе с кодом."""
    
    password = pass_reader() # чтение пароля из файла
    if message.text == password:
        markup = types.ReplyKeyboardRemove()
        sent = bot.send_message(message.chat.id, 'Успешный вход. Загрузите таблицу')
        bot.register_next_step_handler(sent, get_file)
    else:
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        but1 = types.KeyboardButton("Выйти")
        but2 = types.KeyboardButton("Попробовать снова")
        markup.add(but1, but2)
        sent = bot.send_message(message.chat.id, 'Пароль неверен', reply_markup=markup)
        bot.register_next_step_handler(sent, try_again_pass_tab)

def get_file(message): # загрузка обновленной БД
    """Загружает таблицу excel-формата, отправленной администратором в боте, в директорию с кодом и после проверки сообщения на тип и расширение документа конвертирует файл в формат базы данных db. 
        Учитывает возможное существование excel и db файлов в директории с кодом, сохраняет старые версии базы данных и excel-файла под названием backup.db и backup.xlsx, соответственно. Сохраняет новые версии data.db и data.xlsx. После завершения работы функции бот отправляет сообщение об успешной загрузке таблицы.
        
    ------
    Exception
        Если файл не имеет расширение xlsx, то бот ответит, что недопустимый тип файла.
        Если файл не является документом, то бот ответит, что файл не является документом.
    """
    
    if message.content_type == 'document': # проверка на тип отправленного сообщения
      URL = bot.get_file_url(message.document.file_id)  # получаем ссылку на документ
      URL_parts = URL.split(".") # дробим ссылку по точкам, чтобы последним элементом списка было расширение файла

      if URL_parts[-1] == 'xlsx': # если расширение xlsx -- бот загружает таблицу
        folder_path = os.getcwd()
        excel_file = os.path.join(folder_path, 'data.xlsx')
        db_file = os.path.join(folder_path, 'data.db')
        backup_xlsx_file = os.path.join(folder_path, 'backup.xlsx')
        backup_db_file = os.path.join(folder_path, 'backup.db')

        if os.path.exists(excel_file): # если уже в директории существует xlsx файл, то проверяем, есть ли бэкап
           if os.path.exists(backup_xlsx_file): # если бэкап есть, то перезаписываем его из data.xlsx
              os.remove(backup_xlsx_file)
              os.rename(excel_file, backup_xlsx_file)
           else:
              os.rename(excel_file, backup_xlsx_file) # если бэкапа нет, то создаем его из data.xlsx

        if os.path.exists(db_file): # если уже в директории существует db файл, то проверяем, есть ли бэкап
           if os.path.exists(backup_db_file): # если бэкап есть, то перезаписываем его из data.db
              os.remove(backup_db_file)
              os.rename(db_file, backup_db_file)
           else:
              os.rename(db_file, backup_db_file) # если бэкапа нет, то создаем его из data.db

        wget.download(URL, 'data.xlsx') # скачиваем из полученной из сообщения ссылки таблицу

        convertation(message, 'data.xlsx') # конвертируем для работы с помощью pandas-SQL

      else: # если расширение НЕ xlsx -- бот пишет что недопустимый тип файла
        bot.send_message(message.chat.id, "Недопустимый тип файла. Загрузите таблицу в формате .xlsx")
    else:
      bot.send_message(message.chat.id, "Загруженный файл не является документом") # если сообщение не документ - говорим об этом

def convertation(message, file_name): # функция конвертации из xlsx в удобный формат для работы с SQL
    """Конвертирует загруженную таблицу из формата excel в формат db. Все хранится в корневой директории вместе с кодом.
        Для успешной работы с db переименовывает поле (колонку) 'Субъект Российской Федерации' в 'Region'.

    Note
        Все файлы хранятся в корневой директории вместе с кодом.
    
    ------
    Exception
        Если excel-таблица не существует в folder_path, то бот сообщает об этом администратору.
        Если в загруженном excel-файле нет листа с названием "СВОД" (что подразумевает загрузку excel-файла с неподходящим содержанием), то бот не обновляет базу данных
    """
    
    try: # проверка по названию рабочего листа - действительно ли загружается БД, а не левый xlsx файл
        folder_path = os.getcwd() # путь к корневой папке
        excel_file = os.path.join(folder_path, file_name) # добавляем путь к файлу
        if os.path.exists(excel_file): # проверяем, что файл по указанному пути существует
            Empty_Rows_Exterminator(excel_file)
            xlsx = pd.ExcelFile(excel_file) # переводим xlsx в pd формат (DataFrame object)
            sheet_name = 'СВОД'
            db = Table_UnRegister(xlsx.parse(sheet_name))
            db = db.rename(columns={'Субъект Российской Федерации': 'Region'})
            db_file = os.path.join(folder_path, 'data.db') # указываем путь для сохранения нового файла
            con = sl.connect(db_file) # создание подключения к базе данных
            cur = con.cursor() # создание объекта "курсор"
            db.to_sql('INV_Table', con, index = False) # , if_exist = 'replace' # собственно конвертация из DataFrame в SQL
            con.commit() # сохранение изменений
            con.close() # закрытие соединения
            xlsx.close() # закрытие файла, иначе перезапись не работает и код ложится отдыхать
            bot.send_message(message.chat.id, "Таблица загружена. Можете работать с обновленной базой данных") # успех
        else:
            bot.send_message(message.chat.id, "Неизвестная ошибка. Попробуйте снова") # иначе пробуем снова
    except: # если пытаются загрузить xlsx, но не подходящую БД, возвращаем предыдущую версию из бэкапа
       folder_path = os.getcwd()
       excel_file = os.path.join(folder_path, 'data.xlsx')
       backup_xlsx_file = os.path.join(folder_path, 'backup.xlsx')
       db_file = os.path.join(folder_path, 'data.db')
       backup_db_file = os.path.join(folder_path, 'backup.db')
       os.remove(excel_file)
    #    os.remove(db_file)
       os.rename(backup_xlsx_file, excel_file)
       os.rename(backup_db_file, db_file)
       bot.send_message(message.chat.id, "Неподходящий Excel-файл. Загрузите базу данных с информацией о регионах")

def pass_reader():
    """Чтение файла с паролем password.txt. Считывает из файла строку с паролем и возвращает её."""
    
    f = open('password.txt','r')
    str = f.readline()
    return str

def pass_writer(message):
    """Перезаписывает строку в файле с паролем password.txt на введенную администратором."""
    
    if message.text != None:
       password = message.text
       f = open('password.txt','w')
       f.truncate(0)
       f.write(password)
       markup = types.ReplyKeyboardRemove()
       bot.send_message(message.chat.id, 'Пароль изменен!', reply_markup=markup)
    else:
       markup = types.ReplyKeyboardRemove()
       bot.send_message(message.chat.id, 'Отправленное сообщение не может быть опознано как пароль. Введите текстовый пароль', reply_markup=markup)
      

def Empty_Rows_Exterminator(excel_path):
    """Проверяет лист 'СВОД' загруженной excel-таблицы на наличие пустых строк и убирает их функцией row_checker, при наличии. Требуется для корректной работы с базой данных после конвертации."""
    
    if __name__ == '__main__':
        book = op.load_workbook(excel_path)
        sheet = book['СВОД']
        for row in sheet:
            row_checker(sheet,row)
    book.save(excel_path)


def row_checker(sheet, row):
    """Проверяет пустые строки excel-таблицы и удаляет их"""
    
    for cell in row:
        if cell.value != None:
              return
    sheet.delete_rows(row[0].row, 1)





@bot.message_handler(commands=['dataset_info']) #работа с БД
def dataset_info(message):
    """После ввода команды /dataset_info становится возможной работа с базой данных. Предлагает начать работу вводом полного наименования субъекта Российской Федерации"""
    
    markup = types.ReplyKeyboardRemove()
    sent = bot.send_message(message.chat.id, 'Введите полное наименование субъекта Российской Федерации', reply_markup=markup)
    bot.register_next_step_handler(sent, data_work)

def data_work(message):
    """Создает словарь со всей информацией о введенном пользователем субъекте РФ. 
    Если в данном субъекте нет ИНВ, то бот сразу сообщает об этом и предлагает клавиатуру с кнопками "Выйти" и "Ввести другой регион".
    Если в данном субъекте есть ИНВ, то бот предлагает пользователю меню с кнопками для дальнейшего предоставления информации о субъекте

    ------
    Exception
        Если наименование субъекта, который ввел пользователь, отсутствует в базе данных, то бот сообщает об этом и предлагает клавиатуру с кнопками "Выйти" и "Попробовать снова".
    """
    
    markup = types.ReplyKeyboardRemove()
    con = sl.connect('data.db')
    user_message = UnRegister(str(message.text))
    try:
      with con:
        data = con.execute('SELECT * FROM INV_Table WHERE Region = :region_;', {'region_': user_message})
        for row in data:
          # dict.fromkeys(seq[, value]) -- можно попробовать
            dict = {
              'Субъект Российской Федерации' : row[1],
              'Закон ИНВ' : row[2],
              'ИНВ для телекома' : row[3],
              'Максимальный размер вычета' : row[4],
              'Ставка для исчисления предельной величины' : row[5],
              'Участие в программе увеличение производительности труда' : row[6],
              'Инвестпроект' : row[7],
              'ОКВЭД 61' : row[8],
              '2022' : row[9],
              '2023' : row[10],
              '2024' : row[11]
              }
      con.close()

      if dict.get('ИНВ для телекома') == 'НЕТ':
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        but1 = types.KeyboardButton("Выйти")
        but2 = types.KeyboardButton("Ввести другой регион")
        markup.add(but1, but2)
        sent = bot.send_message(message.chat.id, 'Для данного региона нет ИНВ', reply_markup=markup)
        bot.register_next_step_handler(sent, try_again_data)
      else:
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        but1 = types.KeyboardButton("Выйти")
        but3 = types.KeyboardButton("Закон ИНВ")
        but4 = types.KeyboardButton("ИНВ для телекома")
        but5 = types.KeyboardButton("Максимальный размер вычета")
        but6 = types.KeyboardButton("Ставка для исчисления предельной величины")
        but7 = types.KeyboardButton("Участие в программе увеличение производительности труда")
        but8 = types.KeyboardButton("Инвестпроект")
        but9 = types.KeyboardButton("ОКВЭД 61")
        but10 = types.KeyboardButton("2022")
        but11 = types.KeyboardButton("2023")
        but12 = types.KeyboardButton("2024")
        markup.add(but1, but3, but4, but5, but6, but7, but8, but9, but10, but11, but12)
        sent = bot.send_message(message.chat.id, 'Что хотите узнать?', reply_markup=markup)
        bot.register_next_step_handler(sent, inform, dict)
    except:
      markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
      but1 = types.KeyboardButton("Выйти")
      but2 = types.KeyboardButton("Попробовать снова")
      markup.add(but1, but2)
      sent = bot.send_message(message.chat.id, 'Регион с таким названием не найден', reply_markup=markup)
      bot.register_next_step_handler(sent, try_again_data)

def UnRegister(text):
    """Преобразует строку в формат Unicode NFKD, позволяющий проводить сравнение строк без учета регистра"""

    return unicodedata.normalize("NFKD", text.casefold())

def Table_UnRegister(excel_file):
    """Применяет функцию UnRegister к столбцу 'Субъект Российской Федерации' из excel-файла"""
    
    counter=0
    while str(excel_file.loc[counter, 'Субъект Российской Федерации']) != 'nan':
        region = str(excel_file['Субъект Российской Федерации'][counter])
        region_normalize = UnRegister(region)
        excel_file.loc[counter, 'Субъект Российской Федерации'] = region_normalize
        counter=counter+1
    return excel_file



def inform(message, dict):
    """После введения пользователем наименования субъекта РФ при наличии ИНВ в данном субъекте пользователь нажимает на одну из предлагаемых кнопок.
    Функция inform предоставляет пользователю информацию исходя из выбранной им кнопки"""

    if message.text == 'Выйти':
        markup = types.ReplyKeyboardRemove()
        bot.send_message(message.chat.id, "Вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)
    else:
        if dict.get(message.text) == None:
          sent = bot.send_message(message.chat.id, message.text + ':\n' + "Нет информации")
          bot.register_next_step_handler(sent, inform, dict)
        else:
          sent = bot.send_message(message.chat.id, message.text + ':\n' + dict.get(message.text))
          bot.register_next_step_handler(sent, inform, dict)

def try_again_data(message):
    """После команды /dataset_info в боте пользователь может ввести сообщение "Выйти" или выбрать эту кнопку на клавиатуре для возвращения на основную страницу.
        Текстовые сообщения "Попробовать снова" и "Ввести другой регион" возвращают к работе с базой данных."""
    
    markup = types.ReplyKeyboardRemove()
    if message.text == 'Выйти':
        bot.send_message(message.chat.id, "Вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)
    elif message.text == 'Попробовать снова' or message.text == 'Ввести другой регион':
        sent = bot.send_message(message.chat.id, 'Введите полное наименование субъекта Российской Федерации', reply_markup=markup)
        bot.register_next_step_handler(sent, data_work)
    else:
        bot.send_message(message.chat.id, "В любом случае, вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)


def try_again_pass_change(message):
    """После неверно введеного пароля в команде /change_password, пользователь может ввести сообщения "Выйти" и "Попробовать снова" или выбрать эти кнопки на клавиатуре для продолжения работы."""
    
    markup = types.ReplyKeyboardRemove()
    if message.text == 'Выйти':
        bot.send_message(message.chat.id, "Вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)
    elif message.text == 'Попробовать снова':
        sent = bot.send_message(message.chat.id, 'Введите старый пароль', reply_markup=markup)
        bot.register_next_step_handler(sent, pass_login)
    else:
        bot.send_message(message.chat.id, "В любом случае, вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)


def try_again_pass_tab(message):
    """После неверно введеного пароля в команде /upload_excel, пользователь может ввести сообщения "Выйти" и "Попробовать снова" или выбрать эти кнопки на клавиатуре для продолжения работы. """
    markup = types.ReplyKeyboardRemove()
    if message.text == 'Выйти':
        bot.send_message(message.chat.id, "Вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)
    elif message.text == 'Попробовать снова':
        sent = bot.send_message(message.chat.id, 'Введите пароль', reply_markup=markup)
        bot.register_next_step_handler(sent, tab_login)
    else:
        bot.send_message(message.chat.id, "В любом случае, вы вернулись на основную страницу", parse_mode='html', reply_markup=markup)




@bot.message_handler(content_types=all_types) # ответ на простое сообщение
def send_answer(message):
    """Просит ввести текст соответствующей команды на основной странице. Бот сообщает о некорректном сообщении."""
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    bot.reply_to(message, message.chat.first_name + ", ваше сообщение не распознано. Введите команду из меню.", parse_mode='html', reply_markup=markup)

bot.polling(non_stop=True)
