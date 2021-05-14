import base64
import os
import tempfile

from class_email import Email
from class_mssql import MSSQLConnection
from class_settings import Settings

set = Settings('settings.json')
sql = MSSQLConnection(set.param('sql_server'), set.param('sql_database'), set.param('sql_login'),
                      set.param('sql_password'))
mail = Email(sender=set.param('email_sender'), reply_to=set.param('email_reply_to'),
             recipient='a.kovalenko@pokupochka.ru', subject='Тестовое письмо',
             message='<h1>Содержимое сообщения</h1><br><p>тест</p>')
mail.server = set.param('mail_server')
mail.port = set.param('mail_port')
mail.login = set.param('mail_login')
mail.password = set.param('mail_password')

query = sql.select('select document_uid, name, convert(nvarchar(max), convert(varbinary(max), data)) from [dbo].['
                   'document_attachments] as att')
attachmets = {}

for x in query:
    path = tempfile.gettempdir() + '\\' + str(x[0])
    file = path + '\\' + x[1]
    os.makedirs(path, exist_ok=True)
    with open(file, 'wb') as f:
        try:
            f.write(base64.b64decode(x[2], validate=True))
            if attachmets.get(str(x[0])) is None:
                attachmets[str(x[0])] = {'files': []}
            attachmets[str(x[0])]['files'].append({x[1]: file})
        except Exception as E:
            print(f'Ошибка записи данных в файл: {file} - {E}')

# Собрали все вложения, предварительно выгрузив их из БД.
print(attachmets)

if mail.connect_smtp():
    mail.attach_file('c:\Temp\9f5837d1-4fe4-11eb-80f4-4c52620055bf\ОтчетПоПродажам.xlsx')
    mail.attach_file("c:\Temp\9f5837d1-4fe4-11eb-80f4-4c52620055bf\ОтчетДляРетроБонусов.xlsx")
    if mail.send_email():
        print('Mail sent')
