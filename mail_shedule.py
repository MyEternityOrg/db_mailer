import base64
import os
import shutil
import tempfile
import time

from class_email import Email
from class_mssql import MSSQLConnection
from class_settings import Settings

# Создаем все нужные классы.
# Настройки.
set = Settings('settings.json')
# Подключение к БД
sql = MSSQLConnection(set.param('sql_server'), set.param('sql_database'), set.param('sql_login'),
                      set.param('sql_password'))
# Почта.

query = sql.select('select * from get_mail_list()')
mail_list = {}

# Получаем наши будущие письма
for x in query:
    path = tempfile.gettempdir() + '\\' + str(x[0])
    if mail_list.get(str(x[0])) is None:
        mail_list[str(x[0])] = {'recipients': x[1], 'copy_to': x[2], 'blind_copy_to': x[3], 'topic': x[4], 'body': x[5],
                                'reply_to': x[6], 'temp_dir': path, 'mail_uid': x[10], 'files': []}
    if x[7] == 1:
        file = path + '\\' + x[8]
        os.makedirs(path, exist_ok=True)
        with open(file, 'wb') as f:
            try:
                f.write(base64.b64decode(x[9], validate=True))
                mail_list[str(x[0])]['files'].append({x[8]: file})
            except Exception as E:
                print(f'Ошибка записи данных в файл: {file} - {E}')

# Собрали все данные для отправки почты.
if len(mail_list) == 0:
    print('Нет почты для отправки.')

for buff in mail_list.keys():
    time.sleep(3)
    record = mail_list[buff]
    print(f'Отправялем письмо: {record["mail_uid"]}')
    mail = Email(sender=set.param('email_sender'), reply_to=record['reply_to'],
                 recipient=record['recipients'], subject=record['topic'],
                 message=record['body'])
    mail.server = set.param('mail_server')
    mail.port = set.param('mail_port')
    mail.login = set.param('mail_login')
    mail.password = set.param('mail_password')
    for f in record['files']:
        mail.attach_file(*f.values())
    try:
        if mail.connect_smtp():
            if mail.send_email(send_mail=True):
                sql.execute('update [mail_buffer] set processed = 1 where uid = %s', (str(record["mail_uid"])))
                sql.execute('exec [set_next_status] %s, 1', (buff))
                print(f'Письмо отправлено: {record["mail_uid"]}')
    except Exception as E:
        sql.execute('exec [set_next_status] %s, 0', (buff))
        print(f'Ошибка отправки почты: {E}')
    # Приберемся за собой.
    shutil.rmtree(record['temp_dir'], ignore_errors=True)
