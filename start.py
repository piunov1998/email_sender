import smtplib, json, ssl, os, sys, win32com.client
from email.mime.text import MIMEText
from email.header import Header

#SMTP server info
smtp_server = 'smtp.gmail.com'
port = 465

if not os.path.exists('./userdata'):
    os.mkdir('./userdata')
    with open('./userdata/logins.json', 'x', encoding = 'utf-8') as file:
        json.dump(
            {
                'your_email_1@pipka.ru' : 'password',
                'your_email_3@pipka.ru' : 'password',
                'your_email_2@pipka.ru' : 'password'
            }, file, sort_keys = True, indent = 2)
    with open('./userdata/message.txt', 'x', encoding = 'utf-8') as file:
        file.write('Заголовок\nТело сообщения')
    input('Заполните файлы в папке userdata своими данными и перезапустите программу.\nНажмите Enter для завершения...')
    sys.exit()

logins = []
emails = []
receiver_email = ['piunov.doc@yandex.ru']

Excel = win32com.client.Dispatch("Excel.Application")
data = Excel.Workbooks.Open(u'C:\\Users//Тоха//github//email_sender//userdata//data.xlsx')
sheet = data.ActiveSheet

i = 0
e = 0
go = True

while go:
    i += 1
    cell = sheet.Cells(i, 1).value
    if sheet.Cells(i, 1).value == None:
        e += 1
    if '@' in str(cell):
        emails.append(cell)
    if e >= 20:
        go = False
data.Close()
Excel.Quit()

with open('./userdata/logins.json', 'r', encoding = 'utf-8') as file:
    logins = json.load(file)

with open('./userdata/message.txt', 'r', encoding = 'utf-8') as file:
    text = file.read().splitlines()
    head = text[0]
    text.pop(0)
    body = ''
    for line in text:
        body += f'{line}\n'
    message = MIMEText(body[:-1], 'plain', 'utf-8')
    message['Subject'] = Header(head, 'utf-8')

last_email = ''
one_email_count = ''

with open('settings.json', 'r', encoding = 'utf-8') as file:
    settings = json.load(file)
    last_email = settings['last_email']
    one_email_count = settings['one_email_count']

start = last_email
stop = start + one_email_count

for login in logins:
    sender_email = login
    password = logins[login]

    for i in range(start, stop):
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context = context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, emails[i], message.as_string())

input('Дело сделано! Нажмите Enter для выхода...')
