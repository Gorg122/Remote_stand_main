from __future__ import print_function
import httplib2
import os
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import pymysql.cursors
import time
import email
import imaplib

# import zipfile
# from New_all_2 import Launch
# from emailsend import send_email
# import smtplib
# import mimetypes                                            # Импорт класса для обработки неизвестных MIME-типов, базирующихся на расширении файла
# from email import encoders                                  # Импортируем энкодер
# from email.mime.base import MIMEBase                        # Общий тип
# from email.mime.text import MIMEText                        # Текст/HTML
# from email.mime.image import MIMEImage                      # Изображения
# from email.mime.audio import MIMEAudio                      # Аудио
# from email.mime.multipart import MIMEMultipart              # Многокомпонентный объект
import datetime
# import shutil
# import io
# from googleapiclient.http import MediaIoBaseDownload

email_name = "Not_email"

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
#CLIENT_SECRET_FILE = 'client_secret_2.json'
CLIENT_SECRET_FILE = 'client_secret_Petukhov.json'
APPLICATION_NAME = 'Drive API Python Quickstart'

def main():
    #################################### Должно быть в начале 1 раз #############
    credentials = get_credentials()                                 ########  работа с exel таблицей
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    service_sheets = discovery.build('sheets', 'v4', http=http)
    #############################################################################
    # addr = "sasha.lorens@yandex.ru"  # Отправитель \\\\\\\\\\\\\\ Работа с почтой
    # password = "LeNoVo_13572468"
    # mail = imaplib.IMAP4_SSL('imap.yandex.ru')
    # mail.login(addr, password)

    return service_sheets
    #################################### Должно быть в начале 1 раз #############

def read_file(filename):
    with open(filename, 'rb') as f:
        data_zip = f.read()
    return data_zip


def write_file(data, filename):
    with open(filename, 'wb') as f:
        f.write(data)
    return 'OK'


def file_upload(my_id, filename, email):
    con = connect()
    data = read_file(filename)
    with con:
        cur = con.cursor()
        try:
            sql = ("""UPDATE status
                      SET soc_zip = %s, email_name = %s, file_id = 0
                      WHERE id = %s""")
            cur.execute(sql, (data, email, my_id))
            con.commit()
            print("Все нормально")
        except:
            print("Какой то кал")


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    # credential_path = os.path.join(credential_dir,
    #                                'drive-python-quickstart2.json')
    credential_path = os.path.join(credential_dir,
                                     'ul_cad_1.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def connect():
    con = pymysql.connect(host='localhost',
                          user='root',
                          password='root',
                          database='labstandstatus',
                          cursorclass=pymysql.cursors.DictCursor)
    return con

def status_check():  ############# Функция  проверки статуса и выставления ################
    con = connect()
    with con:
        cur = con.cursor()
        cur.execute("SELECT id FROM status WHERE status = 1")
        answer = cur.fetchall()
        try:
            clear_for_work_stand = answer[0]['id']
            print(clear_for_work_stand)
            return clear_for_work_stand
        except:
            print("На данный момент свободных стендов нет")
            return 0

def send_id_for_download(id_dwnld, email_name_sql, clear_for_work_stand ): ##### Функция меняющая id файла который необходимо скачать
    con = connect()
    with con:
        cur = con.cursor()
        # try:
        sql = ("""UPDATE status
                                        SET file_id = %s, email_name = %s
                                        WHERE id = %s""")
        cur.execute(sql, (id_dwnld, email_name_sql,clear_for_work_stand))
        con.commit()
        change_status(clear_for_work_stand, 2)
        print("Все нормально")
        # except:
        # print("Какой то кал")

def change_status(my_id, status_change):
    con = connect()
    with con:
        cur = con.cursor()

        sql = ("""UPDATE status
                 SET status = %s
                 WHERE id = %s""")
        print(status_change, my_id)
        cur.execute(sql, (status_change,my_id))
        con.commit()
    return ('OK')


def exel_work(service_sheets):
    ranges = ["A2:C2"] #в этом месте надо выбрать ечейку которые будем исспользовать.
    #spreadsheetId2 = "1hNTK6F98X5-lB1TIialANY9diKIXrQXRUQKMTVrKzB4"
    spreadsheetId2 = "1hydoacEI1g9zjaLma-NTmLf1OhgeulSbyRSB53M6tXo"
    results = service_sheets.spreadsheets().values().batchGet(spreadsheetId = spreadsheetId2,
                                     ranges = ranges,
                                     valueRenderOption = 'UNFORMATTED_VALUE',
                                     dateTimeRenderOption = 'FORMATTED_STRING').execute()
    try:
        sheet_values = results['valueRanges'][0]['values'][0][2]
        value_id = sheet_values.split('=')[1]
        email_name = results['valueRanges'][0]['values'][0][1]
        print(value_id)
        return (value_id, email_name)
    except:
        print("Значений нет")
        return (0, 0)

def exel_del(service_sheets):
    ranges = ["A2:C2"]  # в этом месте надо выбрать ечейку которые будем исспользовать.
    #spreadsheetId2 = "1hNTK6F98X5-lB1TIialANY9diKIXrQXRUQKMTVrKzB4"
    spreadsheetId2 = "1hydoacEI1g9zjaLma-NTmLf1OhgeulSbyRSB53M6tXo"
    results = service_sheets.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId2,
                                                              ranges=ranges,
                                                              valueRenderOption='UNFORMATTED_VALUE',
                                                              dateTimeRenderOption='FORMATTED_STRING').execute()

    try:
        results_del = service_sheets.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId2, body={
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            #"sheetId": 1843988947,
                            "sheetId": 123680228,
                            "dimension": "ROWS",
                            "startIndex": 1,
                            "endIndex": 2
                        }
                    }
                }
            ]
        })
        results_del.execute()
    except:
        print("Значений нет")
        return (0, 0)


def file_mail_download(mail):
    download_folder = r'C:\Users\grish\PycharmProjects\For_Sasha\Prototype_new_2\Prototype_new_2\download_for_bd'
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        # filename = part.get_filename()
        filename = 'file.zip'
        att_path = os.path.join(download_folder, filename)
        if not os.path.isfile(att_path):
            fp = open(att_path, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()



def mail_find():
    # addr = "sasha.lorens@yandex.ru"  # Отправитель \\\\\\\\\\\\\\ Работа с почтой
    # password = "LeNoVo_13572468"
    addr = "desorder2881488@yandex.ru"  # Отправитель \\\\\\\\\\\\\\ Работа с почтой
    password = "781430DesloG1"
    try:
        mail = imaplib.IMAP4_SSL('imap.yandex.ru')
        mail.login(addr, password)
        mail.list()
        mail.select("inbox")
        mail.select(readonly=False)
        term = u"CAD_MIEM_SOC".encode("utf-8")  ######
        mail.literal = term  ######
        result, data = mail.search("utf-8", "SUBJECT")  ###### ЭТО РАБОТАЕТ!!!!!!!!!!!!!!!!!!!!!!!
        print(result)
        if result == 'OK':
            ids = data[0]
            print(data)
            id_list = ids.split()
            print(id_list)
            latest_email_id = id_list[-1]
            result, data = mail.fetch(latest_email_id, "(RFC822)")
            raw_email = data[0][1]
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)
            email_from_addr = email.utils.parseaddr(email_message['From'])[1]
            file_mail_download(email_message)
            pc_id = status_check()
            print(pc_id)
            filename = r'C:\Users\grish\PycharmProjects\For_Sasha\Prototype_new_2\Prototype_new_2\download_for_bd\file.zip'
            print('u8g8ugi8hhuih')
            if pc_id != 0:
                file_upload(pc_id, filename, email_from_addr)
                change_status(pc_id, 2)
                print('прошла загрузка/////////////////////////////////////////')
                if os.path.isfile(filename):
                    os.remove(filename)
                    print('delited')
                mail.store(latest_email_id, '+FLAGS', '\\DELETED')
                print(latest_email_id)
                mail.expunge()
                return 1
    except:
        print('Нет писем или ошибка подключения')
        return 0


def infinet_check(service_sheets, file_id, email):
    pc_id = status_check()
    print(pc_id)
    if pc_id != 0:
        send_id_for_download(file_id, email, pc_id)
        print("запись успешно записана")
        exel_del(service_sheets)
        return 1
    time.sleep(5)

def sub_main(service_sheets):
    while True:
        id, email = exel_work(service_sheets)
        mail_find() # поиск и загрузка в базу прошивок
        if (id != 0) and (email != 0):
            some_flag = 0
            while some_flag == 0:
                some_flag = infinet_check(service_sheets, id, email)
                print("Нет свободных компов") ##### Здесь должен быть цикл на проверку новых компов
        else:
            print("Нет записей")
            time.sleep(20)


if __name__ == '__main__':
    sheets = main()
    sub_main(sheets)
