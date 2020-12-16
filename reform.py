import win32com.client, sys, os

user_path = input('Укажите путь к базе email-адресов для их извлечения(Enter для стандартного):\n')
if user_path == '':
    data_path = os.path.abspath('./userdata/data.xlsx')
else:
    data_path = user_path[-(len(user_path) - 1):-1]

Excel = win32com.client.Dispatch("Excel.Application")
data_path = os.path.abspath('./userdata/data.xlsx')
data = Excel.Workbooks.Open(data_path)
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

data = Excel.Workbooks.Add()
#data = data = Excel.Workbooks.Open(u'C:\\Users//Тоха//github//email_sender//userdata//data_new.xlsx')
sheet = data.ActiveSheet

i = 0
for email in emails:
    i += 1
    sheet.Cells(i, 1).value = email
new_data_path = os.path.join(os.path.split(data_path)[0], 'data_new.xlsx')
data.SaveAs(new_data_path)
data.Close()
Excel.Quit()

input(f'Записано {len(emails)} уникальных адресов\nНовая база сохранена по пути: {new_data_path}\nНажмите Enter для выхода...')