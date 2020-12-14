import win32com.client, sys

if input('Указанный файл будет перезаписан, продолжить? (да/нет):\n') == 'да':
    print('Работаем...')
else:
    sys.exit()

Excel = win32com.client.Dispatch("Excel.Application")
data = Excel.Workbooks.Open(u'C:\\Users//Тоха//github//email_sender//userdata//data.xlsx')
sheet = data.ActiveSheet

emails = []
i = 0
e = 0
go = True

while go:
    i += 1
    cell = sheet.Cells(i, 1).value
    if sheet.Cells(i, 1).value == None:
        e += 1
    if '@' in str(cell) and not str(cell) in emails:
        emails.append(cell)
    if e >= 19:
        go = False
data.Close()

emails.sort()

book = Excel.Workbooks.Add()
data = data = Excel.Workbooks.Open(u'C:\\Users//Тоха//github//email_sender//userdata//data_new.xlsx')
sheet = data.ActiveSheet

i = 0
for email in emails:
    i += 1
    sheet.Cells(i, 1).value = email
data.Save()
data.Close()
Excel.Quit()

input(f'Записано {len(emails)} уникальных адресов\nНажмите Enter для выхода...')