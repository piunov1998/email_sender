import smtplib, json, ssl, os, sys, win32com.client, time, datetime, shutil
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email import encoders
from email.mime.base import MIMEBase

#SMTP server info
smtp_server = 'smtp.gmail.com'
port = 465

#Message
subject = ''
body = ''
message = MIMEMultipart()

#last_email = ''
#one_email_count = ''
#delay = 0
email_num = 0
logins = []
emails = []

if not os.path.exists('./userdata'):
    os.makedirs('./userdata/attachmets')
    with open('./userdata/logins.json', 'x', encoding = 'utf-8') as file:
        json.dump(
            {
                'your_email_1@pipka.ru' : 'password',
                'your_email_3@pipka.ru' : 'password',
                'your_email_2@pipka.ru' : 'password'
            }, file, sort_keys = True, indent = 2)
    with open('./userdata/message.txt', 'x', encoding = 'utf-8') as file:
        file.write('Заголовок\nТело сообщения')
    with open('./settings.json', 'w', encoding = 'utf-8') as file:
        user_path = input('Укажите путь к базе email-адресов(Enter для стандартного):\n')
        if user_path == '':
            data_path = './userdata/data.xlsx'
        else:
            data_path = user_path[-(len(user_path) - 1):-1]
        json.dump({
            'delay' : int(input('Введите задержу между отправками в секундах:\n')),
            'one_email_count' : int(input('Введите количество адресатов в одном письме:\n')),
            'last_email' : 0,
            'data_path' : data_path
        }, file, sort_keys = True, indent = 2)
    input('Заполните файлы в папке userdata своими данными и перезапустите программу.\nНажмите Enter для завершения...')
    sys.exit()


with open('settings.json', 'r', encoding = 'utf-8') as file:
    settings = json.load(file)
    last_email = settings['last_email']
    one_email_count = settings['one_email_count']
    delay = settings['delay']
    data_path = settings['data_path']

Excel = win32com.client.Dispatch("Excel.Application")
data = Excel.Workbooks.Open(data_path)
shutil.copy(data_path, data_path + '.backup')
sheet = data.ActiveSheet

i = 0
e = 0
go = True
send_break = False

while go:
    i += 1
    cell = sheet.Cells(i, 1).value
    if sheet.Cells(i, 1).value == None:
        e += 1
    if '@' in str(cell):
        emails.append(cell)
    if e >= 20:
        go = False

with open('./userdata/logins.json', 'r', encoding = 'utf-8') as file:
    logins = json.load(file)

with open('./userdata/message.txt', 'r', encoding = 'utf-8') as file:
    text = file.read().splitlines()
    head = text[0]
    text.pop(0)
    body = ''
    for line in text:
        body += f'{line}\n'
    message.attach(MIMEText(body[:-1], 'plain', 'utf-8'))
    message['Subject'] = Header(head, 'utf-8')

for filename in os.listdir('./userdata/attachmets'):
    with open(f'./userdata/attachmets/{filename}', 'rb') as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )
        message.attach(part)

def progress(num):
    string = f'Отправлено {email_num} сообщений, ошибок - 0'
    sys.stdout.write(string)
    sys.stdout.flush()
    sys.stdout.write('\b' * (len(string)))

def conspect(email):
    date = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    for i in range(one_email_count):
        sheet.Cells(last_email + i, 2).value = date
        sheet.Cells(last_email + i, 3).value = email
    data.Save()
    
while not send_break:
    for login in logins:
        sender_email = login
        password = logins[login]
        receivers = []
        for i in range(last_email, last_email + one_email_count):
            try:
                receivers.append(emails[i])
            except:
                break

        last_email += one_email_count

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context = context) as server:
            server.login(sender_email, password)
            try:
                server.sendmail(sender_email, receivers, message.as_string())
            except:
                with open('settings.json', 'r+', encoding = 'utf-8') as file:
                    sett = json.load(file)
                    sett['last_email'] = email_num
                    json.dump(sett, file, sort_keys = True, indent = 2)
                    data.Save()
                    data.Close()
                    Excel.Quit()
                input('Произошла ошибка при отправке.\nНажмите Enter для выхода...')
                sys.exit(1)        
        conspect(sender_email)
        email_num += len(receivers)
        progress(email_num)
        if email_num >= len(emails):
            send_break = True
            break
        time.sleep(delay)

input('Дело сделано! Нажмите Enter для выхода...')