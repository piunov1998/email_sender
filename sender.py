import smtplib, json, ssl, os, sys, win32com.client, time, datetime, shutil
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email import encoders
from email.mime.base import MIMEBase

#SMTP server info
smtp_server = 'smtp.gmail.com'
port = 465

#Emails info containers
logins = []
emails = []

if not os.path.exists('./settings.json'):
    if not input('Введите код активации:\n') == 'R47wzGaYiM1DaN1xz':
        input('Неверный код!\nНажмите Enter для выхода..')
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
    input('Заполните файлы в папке userdata своими данными и перезапустите программу.\nНажмите Enter для завершения...')
    sys.exit()

with open('settings.json', 'r', encoding = 'utf-8') as file:
    settings = json.load(file)
    last_email = settings['last_email']
    one_email_count = settings['one_email_count']
    delay = settings['delay']
    data_path = settings['data_path']

print('Загружаем базу данных..')
Excel = win32com.client.Dispatch("Excel.Application")
data = Excel.Workbooks.Open(data_path)
shutil.copy(data_path, './userdata/data.backup')
print(f'Резервная копия была сохранена по пути: {os.path.abspath("./userdata/data.backup")}.')
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

print(f'Загружено {len(emails)} email-адресов.')

with open('./userdata/logins.json', 'r', encoding = 'utf-8') as file:
    logins = json.load(file)

with open('./userdata/message.txt', 'r', encoding = 'utf-8') as file:
    message = MIMEMultipart()
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
    string = f'Отправлено {last_email} сообщений'
    sys.stdout.write(string)
    sys.stdout.flush()
    sys.stdout.write('\b' * (len(string)))

def conspect(email, pack):
    date = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    for add in range(pack):
        empty = False
        row = last_email + 1 + add
        column = 2
        while not empty:
            if sheet.Cells(row, column).value == None:
                sheet.Cells(row, column).value = date
                sheet.Cells(row, column + 1).value = email
                empty = True
            else:
                column += 2
    data.Save()

def data_save():
    with open('settings.json', 'r', encoding = 'utf-8') as file:
        sett = json.load(file)
    sett['last_email'] = last_email
    with open('settings.json', 'w', encoding = 'utf-8') as file:        
        json.dump(sett, file, sort_keys = True, indent = 2)
    data.Save()
    data.Close()
    Excel.Quit()

#Sending cycle   
print('Начинаем отправку..')
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
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(smtp_server, port, context = context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receivers, message.as_string())        
            conspect(sender_email, len(receivers))
            last_email += len(receivers)
            progress(last_email)
            if last_email >= len(emails):
                send_break = True
                break
            time.sleep(delay)
        except KeyboardInterrupt:
            data_save()
            input('\nПрограмма остановлена. Текущеее состояние записанно.\nНажмите Enter для выхода...')
            sys.exit(1)
        except Exception as error:
            data_save()
            input(f'\nПроизошла ошибка:\n{error}\nДанные были сохранены.\nНажмите Enter для выхода...')
            sys.exit(1)

last_email = 0
data_save()
input('\nДело сделано! Нажмите Enter для выхода...')