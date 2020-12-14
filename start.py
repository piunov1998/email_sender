import smtplib, json, ssl, os, sys
from email.mime.text import MIMEText
from email.header import Header

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

smtp_server = 'smtp.gmail.com'
port = 465

logins = []
receiver_email = ['piunov.doc@yandex.ru']

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

for login in logins:
    sender_email = login
    password = logins[login]

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context = context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

input('All work has been done! Press Enter to exit...')
